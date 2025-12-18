"""
Email Configuration Manager
Stores and retrieves SMTP credentials securely
"""

import json
import base64
import os
from pathlib import Path


class EmailConfig:
    def __init__(self, config_file=".email_config.json"):
        self.config_file = Path(config_file)
        self.config = self.load_config()

    def load_config(self):
        """Load existing configuration"""
        if self.config_file.exists():
            try:
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    # Decode password
                    if 'password' in config:
                        config['password'] = base64.b64decode(config['password'].encode()).decode()
                    return config
            except:
                return {}
        return {}

    def save_config(self, smtp_server, smtp_port, email, password):
        """Save SMTP configuration"""
        config = {
            'smtp_server': smtp_server,
            'smtp_port': smtp_port,
            'email': email,
            'password': base64.b64encode(password.encode()).decode()  # Basic encoding
        }

        with open(self.config_file, 'w') as f:
            json.dump(config, f)

        # Make file read-only for current user
        try:
            os.chmod(self.config_file, 0o600)
        except:
            pass  # Windows might not support this

        print(f"✓ Configuration saved to {self.config_file}")

    def has_config(self):
        """Check if configuration exists"""
        return bool(self.config)

    def get_smtp_server(self):
        return self.config.get('smtp_server')

    def get_smtp_port(self):
        return self.config.get('smtp_port')

    def get_email(self):
        return self.config.get('email')

    def get_password(self):
        return self.config.get('password')

    def clear_config(self):
        """Delete saved configuration"""
        if self.config_file.exists():
            self.config_file.unlink()
            print("✓ Configuration cleared")
        self.config = {}
