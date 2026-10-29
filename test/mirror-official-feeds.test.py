"""Verify failure metadata blocks retained XML and one source cannot suppress others."""
import importlib.util, json, tempfile, sys
from pathlib import Path
sys.dont_write_bytecode = True
spec=importlib.util.spec_from_file_location('mirror', Path(__file__).resolve().parents[1]/'tools/mirror-official-feeds.py')
mirror=importlib.util.module_from_spec(spec);spec.loader.exec_module(mirror)
with tempfile.TemporaryDirectory() as directory:
    root=Path(directory);(root/'feeds').mkdir();seed=root/'seed';seed.mkdir()
    sources=[{'id':'good','file':'official-good.xml','url':'https://www.kslegislature.gov/feeds/rss/bills/'},{'id':'bad','file':'official-bad.xml','url':'https://docs.legis.wisconsin.gov/feed/custom/introduced'}]
    (root/'feeds/official-sources.json').write_text(json.dumps(sources))
    (seed/'good.xml').write_text('<rss><channel><item><title>HB1</title></item></channel></rss>')
    (seed/'bad.xml').write_text('<html>outage</html>')
    (root/'feeds/official-bad.xml').write_text('previous successful bytes')
    mirror.ROOT=root;sys.argv=['mirror','--from-dir',str(seed)];mirror.main()
    status=json.loads((root/'feeds/official-feed-status.json').read_text())
    assert status['good']['ok'] is True and status['good']['items']==1
    assert status['bad']['ok'] is False
    assert (root/'feeds/official-bad.xml').read_text()=='previous successful bytes'
    assert (root/'feeds/official-good.xml').exists()
    for content in [b'<!DOCTYPE rss><rss><channel/></rss>', b'<html/>']:
        try: mirror.validate(content)
        except ValueError: pass
        else: raise AssertionError('Unsupported source accepted')
print('Official mirror: independent sources, XML validation, retained-data failure flags and public allowlist passed.')
