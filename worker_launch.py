import json
import subprocess
from pathlib import Path
from typing import Dict, List, Optional

class WorkerLauncher:
    def __init__(self, config_path: str = 'worker_config.json'):
        self.config_path = Path(config_path)
        self.workers = self._load_config()

    def _load_config(self) -> List[Dict]:
        """Load worker configurations from JSON file."""
        if not self.config_path.exists():
            return []
        with open(self.config_path, 'r') as f:
            return json.load(f)

    def save_config(self, workers: List[Dict]):
        """Save worker configurations to JSON file."""
        with open(self.config_path, 'w') as f:
            json.dump(workers, f, indent=2)

    def load_workers_from_file(self, creds_file: str = 'credentials.json') -> bool:
        """Load workers from a JSON credentials file."""
        try:
            with open(creds_file, 'r') as f:
                config = json.load(f)
                workers = config.get('workers', [])
                for worker in workers:
                    self.add_worker(
                        worker.get('worker_id', worker['ip']),  # Use IP as ID if not specified
                        worker['ip'],
                        worker['username'],
                        worker['password']
                    )
                return True
        except Exception as e:
            print(f"Error loading credentials: {e}")
            return False

    def add_worker(self, worker_id: str, ip: str, username: str, password: str):
        """Add a new worker configuration with a unique ID."""
        workers = self._load_config()
        workers.append({
            'worker_id': worker_id,
            'ip': ip,
            'username': username,
            'password': password,
            'status': 'offline',
            'last_seen': None
        })
        self.save_config(workers)
        print(f"Added worker {worker_id} ({username}@{ip})")
    

    def execute_remote(self, ip: str, username: str, password: str, command: str) -> bool:
        """Execute a command on a remote worker."""
        ps_script = f'''
        $ErrorActionPreference = "Stop"
        try {{
            $pass = ConvertTo-SecureString "{password}" -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential ("{username}", $pass)
            $remoteScript = {{
                # This block runs on the remote machine
                $targetDir = "D:\Trainee\Khushi\PythonProjectLAN_client"
                if (-not (Test-Path -Path $targetDir)) {{
                    Write-Error "Directory does not exist on remote machine: $targetDir"
                    exit 1
                }}
                Set-Location $targetDir
                {command}
            }}
            $result = Invoke-Command -ComputerName {ip} -Credential $cred -ScriptBlock $remoteScript -ErrorAction Stop
            $result
        }}
        catch {{
            Write-Error $_.Exception.Message
            exit 1
        }}
        '''
        
        temp_ps = None
        try:
            with tempfile.NamedTemporaryFile(suffix='.ps1', delete=False, mode='w', encoding='utf-8') as f:
                temp_ps = f.name
                f.write(ps_script)
            
            # Execute the script
            result = subprocess.run(
                ['powershell', '-ExecutionPolicy', 'Bypass', '-File', temp_ps],
                capture_output=True,
                text=True
            )
            
            if result.returncode != 0:
                print(f"Error executing on {ip}: {result.stderr}")
                return False
                
            print(f"Output from {ip}:\n{result.stdout}")
            return True
            
        except Exception as e:
            print(f"Failed to execute on {ip}: {str(e)}")
            return False
        finally:
            if temp_ps and os.path.exists(temp_ps):
                try:
                    os.remove(temp_ps)
                except Exception as e:
                    print(f"Warning: Could not remove temporary file {temp_ps}: {e}")