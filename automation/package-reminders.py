#!/usr/bin/env python3
"""Apply the portable reminder definition to a Microsoft-exported connector package."""
import argparse
import json
from pathlib import Path
from urllib.parse import urlsplit
from uuid import UUID
import zipfile

parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('exported_package', type=Path)
parser.add_argument('output', type=Path)
parser.add_argument('--site', required=True)
parser.add_argument('--audit-list', required=True)
args = parser.parse_args()
url = urlsplit(args.site)
if url.scheme != 'https' or not url.hostname or url.username or url.password or url.query or url.fragment:
    parser.error('--site must be an HTTPS SharePoint site URL without credentials or query parameters')
list_id = str(UUID(args.audit_list))
with zipfile.ZipFile(args.exported_package) as source:
    files = {n: source.read(n) for n in source.namelist()}
paths = [n for n in files if n.endswith('/definition.json')]
if len(paths) != 1:
    parser.error('Export a package containing exactly one non-solution cloud flow')
path = paths[0]
package = json.loads(files[path])
refs = package['properties']['connectionReferences']
connections = {}
for kind in ('sharepointonline', 'office365'):
    matches = [key for key, value in refs.items() if value.get('apiName') == kind]
    if len(matches) != 1:
        parser.error('The source flow must contain one SharePoint and one Office 365 Outlook connection')
    connections['shared_' + kind] = matches[0]
definition = json.loads(Path(__file__).with_name('reminders.definition.json').read_text())
definition['parameters']['SharePointSite']['defaultValue'] = args.site.rstrip('/')
definition['parameters']['AuditListId']['defaultValue'] = list_id

def rebind(value):
    if isinstance(value, dict):
        if 'host' in value and isinstance(value['host'], dict):
            host = value['host']
            if host.get('connectionName') in connections:
                host['connectionName'] = connections[host['connectionName']]
        for item in value.values():
            rebind(item)
    elif isinstance(value, list):
        for item in value:
            rebind(item)

rebind(definition)
package['properties']['definition'] = definition
package['properties']['displayName'] = 'Legislative Router — Scheduled reminders'
files[path] = json.dumps(package, ensure_ascii=False).encode()
args.output.parent.mkdir(parents=True, exist_ok=True)
with zipfile.ZipFile(args.output, 'w', zipfile.ZIP_DEFLATED) as dest:
    for name, contents in files.items():
        dest.writestr(name, contents)
print(args.output)
