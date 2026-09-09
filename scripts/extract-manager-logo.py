"""Reproduce the actual logo asset from the template, without drawing a replacement.
The EMF contains a single STRETCHDIBITS 32-bit bitmap, not vector paths.
"""
import base64,pathlib,struct,zipfile,zlib
source=next(pathlib.Path('reference').glob('*.xlsx'))
with zipfile.ZipFile(source) as archive:data=archive.read('xl/media/image1.emf')
pos=0
while pos<len(data):
 kind,length=struct.unpack_from('<II',data,pos)
 if kind==81:break
 if not length:raise ValueError('Invalid EMF record')
 pos+=length
else:raise ValueError('No embedded bitmap')
record=data[pos:pos+length];info,info_size,bits,bits_size=struct.unpack_from('<4I',record,48)
width,height=struct.unpack_from('<ii',record,info+4)
assert struct.unpack_from('<H',record,info+14)[0]==32
pixels=record[bits:bits+bits_size];rows=[]
for y in reversed(range(height)):
 row=bytearray()
 for x in range(width):
  blue,green,red,_=pixels[(y*width+x)*4:(y*width+x+1)*4]
  row.extend((red,green,blue,0 if (red,green,blue)==(255,255,255) else 255))
 rows.append(b'\0'+row)
def chunk(kind,data):return struct.pack('>I',len(data))+kind+data+struct.pack('>I',zlib.crc32(kind+data))
png=b'\x89PNG\r\n\x1a\n'+chunk(b'IHDR',struct.pack('>IIBBBBB',width,height,8,6,0,0,0))+chunk(b'IDAT',zlib.compress(b''.join(rows)))+chunk(b'IEND',b'')
folder=pathlib.Path('src/comparison/assets');folder.mkdir(exist_ok=True)
(folder/'fly-logo.png').write_bytes(png)
(folder/'fly-logo.ts').write_text('// Extracted losslessly from the template EMF embedded bitmap (143 × 44). White backdrop made transparent.\nexport const flyLogo = "data:image/png;base64,'+base64.b64encode(png).decode()+'";\n')
print(width,height,'source pixels; no invented vector or text logo')
