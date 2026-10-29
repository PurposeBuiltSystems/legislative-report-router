#!/usr/bin/env python3
"""Mirror the reviewed public feed catalog. Never fetch user-entered URLs."""
from pathlib import Path
from datetime import datetime, timezone
import argparse, json, re, urllib.request, xml.etree.ElementTree as ET
from urllib.parse import urlparse
ROOT = Path(__file__).resolve().parents[1]
ALLOWED_HOSTS = {'docs.legis.wisconsin.gov', 'www.kslegislature.gov', 'capitol.texas.gov'}
MAX_BYTES = 4 * 1024 * 1024

def validate(data):
    if len(data) > MAX_BYTES or re.search(br'<!DOCTYPE|<!ENTITY', data, re.I):
        raise ValueError('Oversized feed or unsupported DTD/entity')
    root = ET.fromstring(data)
    if root.tag != 'rss' or root.find('channel') is None:
        raise ValueError('Expected RSS with a channel')
    return len(root.findall('./channel/item'))

def main():
    args = argparse.ArgumentParser()
    args.add_argument('--from-dir', type=Path, help='Validate already downloaded official XML without network access')
    args = args.parse_args()
    sources = json.loads((ROOT / 'feeds/official-sources.json').read_text())
    status_path = ROOT / 'feeds/official-feed-status.json'
    previous = json.loads(status_path.read_text()) if status_path.exists() else {}
    statuses = {}
    for source in sources:
        key, filename = source['id'], source['file']
        assert re.fullmatch(r'[a-z-]+', key) and re.fullmatch(r'official-[a-z-]+\.xml', filename)
        url = urlparse(source['url'])
        assert url.scheme == 'https' and url.hostname in ALLOWED_HOSTS and not url.username and not url.password
        checked = datetime.now(timezone.utc).isoformat()
        try:
            if args.from_dir:
                data = (args.from_dir / (key + '.xml')).read_bytes()
            else:
                request = urllib.request.Request(source['url'], headers={'User-Agent': 'LegislativeReportRouter-public-feed-mirror/1.0'})
                with urllib.request.urlopen(request, timeout=25) as response:
                    final = urlparse(response.url)
                    if final.scheme != 'https' or final.hostname not in ALLOWED_HOSTS:
                        raise ValueError('Unexpected redirect destination')
                    data = response.read(MAX_BYTES + 1)
            count = validate(data)
            target = ROOT / 'feeds' / filename
            temp = target.with_suffix('.xml.tmp')
            temp.write_bytes(data)
            temp.replace(target)
            statuses[key] = {'ok': True, 'checkedAt': checked, 'fetchedAt': checked, 'items': count}
            print(key + ': refreshed (' + str(count) + ' feed items)')
        except Exception as error:
            statuses[key] = {'ok': False, 'checkedAt': checked, 'fetchedAt': previous.get(key, {}).get('fetchedAt'), 'error': str(error)[:240]}
            print(key + ': unavailable; retained prior XML, marked unusable: ' + str(error))
    status_path.write_text(json.dumps(statuses, indent=2) + '\n')

if __name__ == '__main__':
    main()
