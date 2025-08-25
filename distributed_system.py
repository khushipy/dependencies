import json
import subprocess
import threading
import time
import os
import socket
from pathlib import Path
from typing import Dict, List, Optional
from datetime import datetime

# Configuration
CONFIG = {
    'host': '0.0.0.0',  # Listen on all network interfaces
    'port': 5000,       # Default port
    'credentials_file': 'credentials.json'  # Worker credentials
}

class WorkerManager:
    def __init__(self):
        self.config = self._load_config()
        self.workers = self.config.get('workers', [])
        self.is_running = True

    def _load_config(self) -> dict:
        """Load configuration from credentials.json"""
        if not Path(CONFIG['credentials_file']).exists():
            print(f"Warning: {CONFIG['credentials_file']} not found. Creating with default config.")
            self._create_default_credentials()
        with open(CONFIG['credentials_file'], 'r') as f:
            return json.load(f)

    def _load_workers(self) -> List[Dict]:
        """Load workers from credentials.json"""
        config = self._load_config()
        return config.get('workers', [])

    def _save_workers(self):
        """Save workers and configuration back to credentials.json"""
        config = {
            'workers': self.workers
        }
        with open(CONFIG['credentials_file'], 'w') as f:
            json.dump(config, f, indent=4)

    def _create_default_credentials(self):
        """Create default credentials file with local worker"""
        default_config = {
            "workers": [{
                "worker_id": socket.gethostname().lower(),
                "ip": socket.gethostbyname(socket.gethostname()),
                "is_local": True,
                "username": os.getlogin(),
                "password": "",
                "status": "offline",
                "last_seen": None
                # worker_dir will be required when adding a worker
            }]
        }
        os.makedirs(os.path.dirname(CONFIG['credentials_file']), exist_ok=True)
        with open(CONFIG['credentials_file'], 'w') as f:
            json.dump(default_config, f, indent=4)
        print(f"Created {CONFIG['credentials_file']} with default local worker.")
        print("Please edit the password before using!")

    def add_worker(self):
        """Add a new worker through interactive prompt"""
        print("\n=== Add New Worker ===")
        worker = {
            "worker_id": input("Worker ID: ").strip(),
            "ip": input("IP Address: ").strip(),
            "is_local": input("Is this the local machine? (y/n): ").lower() == 'y',
            "username": input("Username (use .\\ for local, DOMAIN\\ for domain): ").strip(),
            "password": input("Password: ").strip(),
            "status": "offline",
            "last_seen": None
        }
        
        # Require worker directory for local workers
        if worker['is_local']:
            while True:
                worker_dir = input("Worker directory (required): ").strip()
                if worker_dir:
                    worker['worker_dir'] = worker_dir
                    break
                print("Error: Worker directory is required for local workers")
        self.workers.append(worker)
        self._save_workers()
        print(f"Added worker: {worker['worker_id']}")

    def _map_network_drive(self, share_path: str, username: str, password: str) -> bool:
        """Map a network drive and return success status"""
        try:
            # Clean up any existing connections
            os.system(f'net use {share_path} /delete /y >nul 2>&1')
            
            # Create new connection
            cmd = f'net use {share_path} /user:{username} {password} /persistent:no'
            result = os.system(cmd)
            return result == 0
        except Exception as e:
            print(f"Error mapping network drive: {e}")
            return False

    def _start_local_worker(self, worker: Dict) -> bool:
        """Start a worker on the local machine"""
        try:
            worker_dir = worker['worker_dir']
            worker_exe = os.path.join(worker_dir, 'main.exe')
            
            if not os.path.exists(worker_exe):
                print(f"Error: Worker executable not found at {worker_exe}")
                return False
                
            subprocess.Popen(
                [worker_exe],
                cwd=worker_dir,
                shell=True
            )
            print(f"Started local worker: {worker['worker_id']}")
            worker['status'] = 'running'
            self._save_workers()
            return True
            
        except Exception as e:
            print(f"Error starting local worker: {e}")
            worker['status'] = 'error'
            self._save_workers()
            return False

    def _start_remote_worker(self, worker: Dict) -> bool:
        """Start a worker on a remote machine"""
        try:
            worker_dir = worker['worker_dir']
            
            # First try direct execution if path is accessible
            remote_exe = os.path.join(worker_dir, 'main.exe')
            if os.path.exists(remote_exe):
                subprocess.Popen(
                    [remote_exe],
                    cwd=worker_dir,
                    shell=True
                )
                print(f"Started remote worker directly: {worker['worker_id']}")
                worker['status'] = 'running'
                self._save_workers()
                return True
                
            # If direct access fails, try mapping network drive
            if '\\' in worker_dir:
                # Extract share path (\\computer\share)
                parts = [p for p in worker_dir.split('\\') if p]
                if len(parts) >= 2:
                    share_path = f"\\\\{parts[0]}\\{parts[1]}"
                    
                    if self._map_network_drive(share_path, worker['username'], worker['password']):
                        if os.path.exists(remote_exe):
                            subprocess.Popen(
                                [remote_exe],
                                cwd=worker_dir,
                                shell=True
                            )
                            print(f"Started remote worker using network share: {worker['worker_id']}")
                            worker['status'] = 'running'
                            self._save_workers()
                            return True
                        else:
                            print(f"Error: Worker executable not found at {remote_exe}")
                    else:
                        print(f"Failed to map network share {share_path}. Check credentials and share permissions.")
                else:
                    print("Invalid network path format. Should be \\\\computer\\share")
            
            return False
            
        except Exception as e:
            print(f"Error starting remote worker: {e}")
            worker['status'] = 'error'
            self._save_workers()
            return False

    def start_worker(self, worker: Dict) -> bool:
        """Start a worker process (local or remote)"""
        try:
            # Validate worker configuration
            if 'worker_dir' not in worker or not worker['worker_dir']:
                print(f"Error: No worker directory specified for worker {worker.get('worker_id', 'unknown')}")
                return False
                
            # Start appropriate worker type
            if worker.get('is_local', False):
                return self._start_local_worker(worker)
            else:
                if not worker.get('username') or not worker.get('password'):
                    print(f"Error: Username and password are required for remote worker {worker.get('worker_id', 'unknown')}")
                    return False
                return self._start_remote_worker(worker)
                
        except Exception as e:
            print(f"Unexpected error starting worker: {e}")
            worker['status'] = 'error'
            self._save_workers()
            return False

    def check_worker_status(self, worker: Dict) -> str:
        """Check if a worker is responsive"""
        try:
            if worker['is_local']:
                # Check if main.exe is running
                result = subprocess.run(
                    ['tasklist', '/FI', 'IMAGENAME eq main.exe'],
                    capture_output=True,
                    text=True
                )
                if "main.exe" in result.stdout:
                    worker['status'] = 'running'
                else:
                    worker['status'] = 'stopped'
            else:
                # For remote workers, implement ping or other check
                response = os.system(f"ping -n 1 {worker['ip']} >nul")
                if response == 0:
                    worker['status'] = 'online'
                else:
                    worker['status'] = 'offline'
            
            worker['last_seen'] = datetime.now().isoformat()
            self._save_workers()
            return worker['status']
        except:
            worker['status'] = 'error'
            self._save_workers()
            return 'error'

    def list_workers(self):
        """List all workers with status"""
        print("\n=== Worker Status ===")
        print(f"{'ID':<15} {'IP':<15} {'Type':<8} {'Status':<10} {'Last Seen'}")
        print("-" * 60)
        
        for worker in self.workers:
            status = self.check_worker_status(worker)
            last_seen = worker.get('last_seen', 'never')
            if last_seen:
                try:
                    last_seen = datetime.fromisoformat(last_seen).strftime('%Y-%m-%d %H:%M:%S')
                except:
                    pass
            print(f"{worker['worker_id']:<15} {worker['ip']:<15} "
                  f"{'Local' if worker['is_local'] else 'Remote':<8} "
                  f"{status:<10} {last_seen}")

def main():
    # Initialize worker manager
    manager = WorkerManager()
    
    # Check if any workers exist
    if not manager.workers:
        print("No workers configured. Please add a worker first.")
        return
    
    while True:
        print("\n=== Distributed System Manager ===")
        print("1. List Workers and Status")
        print("2. Start All Workers")
        print("3. Start Specific Worker")
        print("4. Add New Worker")
        print("5. Refresh Worker List")
        print("6. Exit")
        
        choice = input("\nSelect an option (1-6): ").strip()
        
        if choice == '1':
            manager.list_workers()
        elif choice == '2':
            print("\nStarting all workers...")
            for worker in manager.workers:
                print(f"Starting {worker['worker_id']}...")
                manager.start_worker(worker)
        elif choice == '3':
            manager.list_workers()
            try:
                worker_id = input("\nEnter Worker ID to start: ").strip()
                worker = next(w for w in manager.workers if w['worker_id'] == worker_id)
                manager.start_worker(worker)
            except StopIteration:
                print("Worker not found!")
            except Exception as e:
                print(f"Error: {e}")
        elif choice == '4':
            manager.add_worker()
        elif choice == '5':
            manager.workers = manager._load_workers()
            print("Worker list refreshed")
        elif choice == '6':
            print("Exiting...")
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()