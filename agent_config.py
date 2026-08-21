"""
Config reader - reads from same config.json as the server.
Shared configuration between CLI agent and task pane.
"""
import json
import os
import sys

CONFIG_DIR = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'ExcelAgent')
CONFIG_FILE = os.path.join(CONFIG_DIR, 'config.json')

DEFAULT_CONFIG = {
    'apiKey': '',
    'apiEndpoint': '',
    'model': '',
    'autoRun': True,
    'chatHistory': []
}


def read_config():
    """Read config from disk (shared with server)."""
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            saved = json.load(f)
            config = DEFAULT_CONFIG.copy()
            config.update(saved)
            return config
    except (FileNotFoundError, json.JSONDecodeError):
        return DEFAULT_CONFIG.copy()


def write_config(partial):
    """Merge partial config into existing config and save."""
    config = read_config()
    config.update(partial)
    os.makedirs(CONFIG_DIR, exist_ok=True)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2)
    return config


def normalize_endpoint(ep):
    """Normalize API endpoint to include /v1/chat/completions."""
    if not ep:
        return ep
    ep = ep.rstrip('/')
    if not ep.endswith('/chat/completions'):
        if ep.endswith('/v1'):
            ep = ep + '/chat/completions'
        elif not ep.endswith(('/v1/chat/completions', '/chat/completions')):
            ep = ep + '/v1/chat/completions'
    return ep
