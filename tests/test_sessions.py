import io
import json
import os
import subprocess
import tempfile
import unittest
import uuid
import zipfile
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

import app as benef


def workbook(headers, rows, sheet='Planilha1'):
    book = Workbook()
    book.active.title = sheet
    book.active.append(headers)
    for row in rows:
        book.active.append(row)
    stream = io.BytesIO()
    book.save(stream)
    stream.seek(0)
    return stream


def main_workbook():
    return workbook(
        ['Nome', 'CPF', 'E-mail', 'Empresa', 'CNPJ', 'Vínculo', 'Matrícula', 'Admissão'],
        [
            ['Pessoa CLT', '10000000001', 'clt@example.invalid', 'Gauge', '12345678000100', 'CLT', '10', '01/02/2020'],
            ['Pessoa estagiária', '10000000002', 'estagio@example.invalid', 'Gauge', '12345678000100', 'Estagiário', '20', '01/02/2020'],
            ['Pessoa PJ', '10000000003', 'pj@example.invalid', 'Gauge', '12345678000100', 'PJ', '30', '01/02/2020'],
        ],
    )


class SessionSecurityTest(unittest.TestCase):
    def setUp(self):
        folder = tempfile.TemporaryDirectory()
        self.addCleanup(folder.cleanup)
        self.root = Path(folder.name) / 'sessions'
        self.root.mkdir()
        temp_dir = patch.object(benef, 'TEMP_DIR', str(self.root))
        temp_dir.start()
        self.addCleanup(temp_dir.stop)
        self.client = benef.app.test_client()

    def upload(self, endpoint='/api/upload', stream=None, client=None, **fields):
        if stream is None:
            stream = main_workbook()
        return (client or self.client).post(endpoint, data={
            'file': (stream, 'teste.xlsx'), **fields,
        })

    def assert_json_status(self, response, status):
        self.assertEqual(response.status_code, status, response.get_data(as_text=True))
        return response.get_json()

    def assert_download(self, sid, file_type):
        with self.client.get(f'/api/download/{sid}/{file_type}') as response:
            self.assertEqual(response.status_code, 200, response.status)
            self.assertIn('attachment', response.headers['Content-Disposition'])
            self.assertEqual(response.headers['Cache-Control'], 'no-store')
            return response.get_data()

    def test_main_flow(self):
        sid = self.assert_json_status(self.upload(), 200)['session_id']
        data = {'session_id': sid, 'sheet_map': {'Grupo de teste': 'Planilha1'}}
        columns = self.assert_json_status(self.client.post('/api/get-columns', json=data), 200)
        self.assertIn('CPF', columns['sheet_data']['Grupo de teste']['headers'])
        preview = self.assert_json_status(self.client.post('/api/preview', json=data), 200)
        self.assertEqual(preview['total'], 3)
        self.assertEqual(preview['rows'][0]['nome'], 'Pessoa CLT')
        process = self.assert_json_status(self.client.post('/api/process', json=data), 200)
        self.assertEqual([process[k]['eligible'] for k in ('tp', 'wh', 'nv')], [1, 2, 2])
        for file_type in ('totalpass', 'totalpass_csv', 'wellhub', 'wellhub_csv_zip',
                          'newvalue', 'newvalue_csv', 'auditoria', 'auditppt'):
            with self.subTest(file_type=file_type):
                self.assertTrue(self.assert_download(sid, file_type))
        with self.client.post(f'/api/wellhub-zip/{sid}',
                              json={'cnpjs': ['12.345.678/0001-00']}) as response:
            self.assertEqual(response.status_code, 200, response.status)
            with zipfile.ZipFile(io.BytesIO(response.get_data())) as archive:
                self.assertEqual(len(archive.namelist()), 1)
                self.assertTrue(archive.namelist()[0].endswith('_Wellhub.xlsx'))

    def test_mar_saude_flow(self):
        sheet = workbook(
            ['Nome', 'CPF', 'Nome Sindicato'],
            [['Elegível', '10000000001', 'Sindicato'], ['PJ', '10000000002', 'Somente PJ']],
        )
        sid = self.assert_json_status(self.upload('/api/upload-mar-saude', sheet), 200)['session_id']
        data = {'session_id': sid, 'sheet_name': 'Planilha1'}
        self.assert_json_status(self.client.post('/api/get-columns-mar-saude', json=data), 200)
        result = self.assert_json_status(self.client.post('/api/process-mar-saude', json=data), 200)
        self.assertEqual((result['eligible'], result['excluded']), (1, 1))
        for file_type in ('ms_excel', 'ms_csv', 'ms_audit'):
            self.assertTrue(self.assert_download(sid, file_type))

    def test_expert_two_uploads_and_downloads(self):
        headers = ['Nome', 'CPF', 'Desc. Tipo de Vínculo', 'Data de desligamento']
        active = workbook(headers, [
            ['Pessoa CLT', '10000000001', 'CLT', ''],
            ['Pessoa estagiária', '10000000002', 'Estagiário', ''],
        ])
        sid = self.assert_json_status(self.upload('/api/upload-expert', active), 200)['session_id']
        inactive = workbook(headers, [['Pessoa desligada', '10000000003', 'CLT', '01/02/2020']])
        second = self.upload('/api/upload-expert', inactive, session_id=sid, source_type='desligados')
        self.assertEqual(self.assert_json_status(second, 200)['session_id'], sid)
        data = {'session_id': sid, 'sheet_name': 'Planilha1'}
        self.assert_json_status(self.client.post('/api/get-columns-expert', json=data), 200)
        result = self.assert_json_status(self.client.post('/api/process-expert', json=data), 200)
        self.assertEqual((result['ativos'], result['inativos'], result['excluidos']), (1, 1, 1))
        for file_type in ('expert_ativos_excel', 'expert_inativos_excel', 'expert_ativos_csv',
                          'expert_inativos_csv', 'expert_ppt'):
            self.assertTrue(self.assert_download(sid, file_type))

    def test_failed_reupload_preserves_previous_base(self):
        sid = self.assert_json_status(self.upload(), 200)['session_id']
        self.assert_json_status(self.upload('/api/upload-previous', session_id=sid), 200)
        source = self.root / sid / 'previous_source.xlsx'
        contents = source.read_bytes()
        with self.assertLogs(benef.app.logger, 'ERROR'):
            rejected = self.upload('/api/upload-previous', io.BytesIO(b'invalid'), session_id=sid)
        self.assert_json_status(rejected, 500)
        self.assertEqual(source.read_bytes(), contents)
        self.assertEqual(sorted(p.name for p in source.parent.iterdir()),
                         ['.owner', 'previous_client.txt', 'previous_source.xlsx', 'source.xlsx'])

    def test_every_session_route_rejects_another_browser(self):
        sid = self.assert_json_status(self.upload(), 200)['session_id']
        other = benef.app.test_client()
        other.get('/').close()
        for rule in benef.app.url_map.iter_rules():
            with self.subTest(route=rule.rule):
                if '<session_id>' in rule.rule:
                    url = rule.rule.replace('<session_id>', sid).replace('<file_type>', 'totalpass')
                    response = (other.post(url, json={'cnpjs': []}) if 'POST' in rule.methods
                                else other.get(url))
                elif rule.endpoint in {'api_upload_previous', 'api_upload_expert'}:
                    response = self.upload(rule.rule, client=other, session_id=sid)
                elif rule.rule.startswith(('/api/get-columns', '/api/process', '/api/preview')):
                    response = other.post(rule.rule, json={'session_id': sid})
                else:
                    continue
                self.assert_json_status(response, 404)
                response.close()
        self.assertEqual(sorted(p.name for p in (self.root / sid).iterdir()),
                         ['.owner', 'source.xlsx'])

    def test_malformed_sessions_and_legacy_directory(self):
        for value in (None, 123, {}, '../outside', r'..\outside', r'C:\outside', '/tmp/outside'):
            with self.subTest(session_id=value):
                self.assert_json_status(self.client.post('/api/preview', json={'session_id': value}), 400)
        response = self.upload('/api/upload-expert', session_id='../outside')
        self.assert_json_status(response, 400)
        self.assertFalse((self.root.parent / 'outside').exists())
        self.assert_json_status(self.client.get('/api/download/../totalpass'), 400)
        legacy = str(uuid.uuid4())
        legacy_dir = self.root / legacy
        legacy_dir.mkdir()
        (legacy_dir / 'totalpass.csv').write_text('dado legado', encoding='utf-8')
        self.client.get('/').close()
        self.assert_json_status(self.client.get(f'/api/download/{legacy}/totalpass_csv'), 404)

    def test_rejects_session_directory_link(self):
        sid = self.assert_json_status(self.upload(), 200)['session_id']
        session_dir = self.root / sid
        outside = self.root.parent / 'moved-session'
        session_dir.rename(outside)
        if os.name == 'nt':
            result = subprocess.run(['cmd.exe', '/c', 'mklink', '/J', sid, str(outside)],
                                    cwd=self.root, capture_output=True)
            if result.returncode:
                self.skipTest('O ambiente não permite criar junções.')
        else:
            session_dir.symlink_to(outside, target_is_directory=True)
        self.assert_json_status(self.client.post('/api/preview', json={'session_id': sid}), 404)
        self.assert_json_status(self.upload('/api/upload-expert', session_id=sid), 404)
        self.assertEqual(sorted(path.name for path in outside.iterdir()), ['.owner', 'source.xlsx'])

    def test_wellhub_zip_ignores_paths_outside_the_session(self):
        sid = self.assert_json_status(self.upload(), 200)['session_id']
        outside = self.root.parent / 'outside.txt'
        outside.write_text('privado', encoding='utf-8')
        (self.root / sid / 'wh_meta.json').write_text(json.dumps([{
            'cnpj': '1', 'filepath': str(outside), 'filename': 'outside.xlsx',
            'csv_path': str(outside), 'csv_name': 'outside.csv',
        }]), encoding='utf-8')
        for endpoint in ('wellhub-zip', 'wellhub-csv-zip-filtered'):
            with self.client.post(f'/api/{endpoint}/{sid}', json={'cnpjs': ['1']}) as response:
                self.assertEqual(response.status_code, 200, response.status)
                with zipfile.ZipFile(io.BytesIO(response.get_data())) as archive:
                    self.assertEqual(archive.namelist(), [])

    def test_errors_do_not_expose_exceptions(self):
        sid = self.assert_json_status(self.upload(), 200)['session_id']
        with self.assertLogs(benef.app.logger, 'ERROR') as logs:
            with patch.object(benef, 'excel_file_fast', side_effect=ValueError('private employee data')):
                response = self.client.post('/api/preview', json={'session_id': sid})
        result = self.assert_json_status(response, 500)
        self.assertEqual(set(result), {'error'})
        self.assertNotIn('private employee data', response.get_data(as_text=True))
        self.assertNotIn('private employee data', '\n'.join(logs.output))
        with self.assertLogs(benef.app.logger, 'ERROR'):
            with patch.object(benef, '_new_session', side_effect=OSError('secret path')):
                response = self.upload()
        result = self.assert_json_status(response, 500)
        self.assertEqual(set(result), {'error'})
        self.assertNotIn('secret path', response.get_data(as_text=True))
        self.assert_json_status(self.client.post('/api/preview', json=[]), 400)

    def test_local_host_origin_and_cookie(self):
        with self.client.get('/') as response:
            self.assertEqual(response.status_code, 200)
            self.assertIn('HttpOnly', response.headers['Set-Cookie'])
            self.assertIn('SameSite=Strict', response.headers['Set-Cookie'])
            self.assertEqual(response.headers['Cache-Control'], 'no-store')
        self.assert_json_status(self.client.get('/', headers={'Host': 'untrusted.invalid'}), 400)
        self.assert_json_status(self.client.get('/', environ_overrides={'REMOTE_ADDR': '192.0.2.2'}), 403)
        self.assertEqual(self.client.get('/', base_url='http://127.0.0.1:5050').status_code, 200)
        self.assertEqual(self.client.get('/', base_url='http://[::1]:5050').status_code, 200)
        self.assert_json_status(self.client.post('/api/upload', headers={
            'Origin': 'http://localhost:9999',
        }), 403)
        self.assert_json_status(self.client.post('/api/upload', headers={
            'Origin': 'http://localhost',
        }), 400)


if __name__ == '__main__':
    unittest.main()
