"""Exercise the deployment script without building or contacting Azure."""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest

SCRIPT = Path(__file__).resolve().parents[2] / 'deploy.sh'


class DeploymentTests(unittest.TestCase):
    def run_deploy(self, scenario):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copy2(SCRIPT, root / 'deploy.sh')
            (root / 'ConvertLearnToDocWeb.PublishSettings').write_text(
                '<publishData><publishProfile publishMethod="MSDeploy" '
                'publishUrl="convertlearntodocweb.scm.azurewebsites.net:443" '
                'userName="test" userPWD="fake-password" /></publishData>')
            bin_dir = root / 'bin'
            bin_dir.mkdir()
            stubs = {
                'dotnet': '#!/bin/bash\nwhile [[ "$1" != "-o" ]]; do shift; done\nmkdir -p "$2"\necho published > "$2/app.dll"\n',
                'sleep': '#!/bin/bash\nexit 0\n',
                'curl': '''#!/usr/bin/env python3
import json, os, pathlib, sys
args = sys.argv[1:]
scenario = os.environ['SCENARIO']
if '/api/zipdeploy' in args[-1]:
    if '-D' in args:
        host = 'untrusted.example' if scenario == 'wrong_host' else 'convertlearntodocweb.scm.azurewebsites.net'
        pathlib.Path(args[args.index('-D')+1]).write_text('HTTP/1.1 202 Accepted\\r\\nLocation: https://' + host + '/api/deployments/test-id\\r\\n\\r\\n')
    if '-w' in args: print('202', end='')
elif args[-1].endswith('/log'):
    print(json.dumps([{'message':'Deployment Failed: util.promisify is not a function'}]))
else:
    print(json.dumps({'id':'test-id', 'status':3 if scenario == 'failed' else 4, 'complete':True}))
''',
            }
            for name, contents in stubs.items():
                target = bin_dir / name
                target.write_text(contents)
                target.chmod(0o755)
            return subprocess.run(['bash', str(root / 'deploy.sh')], cwd=root,
                                  env={**os.environ, 'PATH': str(bin_dir) + ':' + os.environ['PATH'],
                                       'SCENARIO': scenario}, capture_output=True, text=True, timeout=20)

    def test_accepted_upload_with_failed_deployment_exits_nonzero(self):
        result = self.run_deploy('failed')
        self.assertNotEqual(result.returncode, 0, result.stdout)
        self.assertIn('Deployment failed', result.stdout + result.stderr)

    def test_success_is_confirmed(self):
        result = self.run_deploy('success')
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn('Deployment succeeded', result.stdout)

    def test_status_redirect_cannot_send_credentials_to_another_host(self):
        result = self.run_deploy('wrong_host')
        self.assertNotEqual(result.returncode, 0, result.stdout)
        self.assertIn('Unexpected deployment status URL', result.stdout + result.stderr)


if __name__ == '__main__':
    unittest.main()
