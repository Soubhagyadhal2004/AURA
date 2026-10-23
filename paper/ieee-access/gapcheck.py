import sys,subprocess,glob,os
from PIL import Image
import numpy as np
pdf=sys.argv[1]
for f in glob.glob('/tmp/build_access/gc-*.png'):os.remove(f)
subprocess.run(['pdftoppm','-png','-r','40',pdf,'/tmp/build_access/gc'])
for f in sorted(glob.glob('/tmp/build_access/gc-*.png')):
    im=np.array(Image.open(f).convert('L'));h,w=im.shape
    top=int(h*0.07);bot=int(h*0.93)  # body area
    out=[]
    for name,(a,b) in (('L',(0,w//2)),('R',(w//2,w))):
        sub=im[top:bot,a:b]
        rows=np.where((sub<200).sum(axis=1)>0)[0]
        last=rows.max() if len(rows) else 0
        out.append('%s=%d%%'%(name,100*last/(bot-top)))
    print(os.path.basename(f),' '.join(out))
