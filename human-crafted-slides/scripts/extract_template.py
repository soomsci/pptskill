#!/usr/bin/env python3
"""Read a local PPTX and inventory original objects without flattening groups.
Only embedded media are copied; fonts and external links are not fetched.
Usage: python extract_template.py source.pptx output_directory
"""
import argparse, hashlib, json, posixpath, re
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def tree(e):
    out={'tag':e.tag,'attributes':dict(e.attrib)}
    if e.text and e.text.strip():out['text']=e.text
    if len(e):out['children']=[tree(c) for c in e]
    return out

def extract(source,destination):
    source=Path(source);destination=Path(destination);destination.mkdir(parents=True,exist_ok=True)
    report={'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),'scope':'Local PPTX export only. Export absence does not prove Canva absence. Text is template sample content, not user facts. Transforms preserve nested local coordinate spaces; compose parent transforms before positioning.','slides':[],'media':[]}
    with ZipFile(source) as z:
        p=ET.fromstring(z.read('ppt/presentation.xml'));report['slide_size_emu']=p.find('p:sldSz',NS).attrib
        names=sorted([n for n in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+.xml',n)],key=lambda x:int(re.search(r'slide(\d+)',x)[1]))
        for n in names:
            doc=ET.fromstring(z.read(n));rels={};rp=posixpath.join(posixpath.dirname(n),'_rels',posixpath.basename(n)+'.rels')
            if rp in z.namelist():
                for rel in ET.fromstring(z.read(rp)):
                    a=dict(rel.attrib)
                    if a.get('TargetMode')!='External':a['resolved_part']=posixpath.normpath(posixpath.join(posixpath.dirname(n),a['Target']))
                    rels[a['Id']]=a
            objects=doc.findall('.//p:sp',NS)+doc.findall('.//p:pic',NS)+doc.findall('.//p:graphicFrame',NS)+doc.findall('.//p:cxnSp',NS)
            slide={'part':n,'object_count':len(objects),'group_count':len(doc.findall('.//p:grpSp',NS)),'fonts':sorted({e.get('typeface') for e in doc.iter() if e.get('typeface')}),'colors':sorted({e.get('val') for e in doc.findall('.//a:srgbClr',NS)}),'transition':None,'timing':None,'relationships':rels,'object_tree':tree(doc.find('p:cSld',NS))}
            for key in ['transition','timing']:
                e=doc.find('p:'+key,NS)
                if e is not None:slide[key]=tree(e)
            report['slides'].append(slide)
        for n in z.namelist():
            if not n.startswith('ppt/media/') or n.endswith('/'):continue
            data=z.read(n);dest=destination/'media'/Path(n).name;dest.parent.mkdir(exist_ok=True);dest.write_bytes(data)
            report['media'].append({'part':n,'local_path':'media/'+dest.name,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()})
        report['native_chart_parts']=[n for n in z.namelist() if re.search(r'/charts/[^/]+\.xml$',n)]
    # Split object trees so routine execution reads only a requested slide.
    sd=destination/'slides';sd.mkdir(exist_ok=True)
    for i,slide in enumerate(report['slides'],1):
        detail=sd/f'slide-{i:02}.json'
        detail.write_text(json.dumps(slide,ensure_ascii=False,indent=2)+'\n')
        slide['detail_path']=f'slides/{detail.name}'
        del slide['object_tree']
        del slide['relationships']
    (destination/'inventory.json').write_text(json.dumps(report,ensure_ascii=False,indent=2)+'\n')
    return {'slides':len(report['slides']),'objects':sum(s['object_count'] for s in report['slides']),'groups':sum(s['group_count'] for s in report['slides']),'media':len(report['media']),'timing_slides':[i+1 for i,s in enumerate(report['slides']) if s['timing']],'transition_slides':[i+1 for i,s in enumerate(report['slides']) if s['transition']]}
if __name__=='__main__':
    ap=argparse.ArgumentParser();ap.add_argument('source');ap.add_argument('destination');args=ap.parse_args();print(json.dumps(extract(args.source,args.destination),ensure_ascii=False))
