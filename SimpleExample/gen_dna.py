# MaZhaoxin @20230507
# Usage:
#   1. Copy this script to *.sln located level.
#   2. Run it.
#   3. Find *.xll, *.dna and auto_load.vbs in `./Distribution`.
# Note:
#   1. Must involve Excel namespace and rename `Excel.Application` to `XlApp`.
#   2. Must delare `app`, `wb` and `ws` for `Excel.Application`, `Workbook` and `Worksheet`.
#   3. Must init them in constructor by
#       app = new XlApp();
#       app.Visible = true;
#       wb = (Workbook)app.Workbooks.Add();
#       ws = (Worksheet)wb.Worksheets.Add();


XLL_PATH = r'E:\Programming\ExcelDNA\ExcelDna64.xll'
OUT_DIR = 'Distribution'

#==================================================
import os, sys, shutil, re

cwd = os.path.abspath(os.curdir)
proj_name = cwd.split('\\')[-1]


def exist_or_error(name):
    if not os.path.exists(name):
        print(f'*E: Cannot find {name}.')
        sys.exit()


exist_or_error(proj_name+'.sln')
exist_or_error(proj_name)
exist_or_error(proj_name+'/MainForm.cs')

#==================================================
DNA_TEMPLATE = '''<DnaLibrary RuntimeVersion="v4.0" Language="C#" >
__REFERENCE__
<![CDATA[
__CSHARP__
]]>
<CustomUI>
__XML__
</CustomUI>
</DnaLibrary>'''

REF_TEMPLATE = '<Reference Name="__REF__" />'

XML_TEMPLATE = '''<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui' onLoad='RibbonLoad'>
  <ribbon>
    <tabs>
      <tab id='CustomTab' label='Custom Tab'>
        <group id='SampleGroup' label='Sample'>
          <button id='Button1' label='Test-1' imageMso='M' size='large' onAction='Button1Click' />
          <button id='Button2' label='Test-2' imageMso='M' size='large' onAction='Button2Click' />
        </group >
      </tab>
    </tabs>
  </ribbon>
</customUI>'''

VBS = '''
'''

# Read files --------------------------------------
if not os.path.exists(proj_name+'.xml'):
    with open(proj_name+'.xml', 'w', encoding='utf8') as f:
        f.write(XML_TEMPLATE)
    xml = XML_TEMPLATE
else:
    xml = open(proj_name+'.xml', encoding='utf8').read()

actions = re.findall(r' on\S+=(\S+)', xml)
actions = [a[1:-1] for a in actions] # remove ' or "

csharp = open(proj_name+'/MainForm.cs', encoding='utf8').read()

ref = []
for r in re.findall(r'using (\S+);', csharp):
    if r not in ['System', 'System.Collections.Generic']:
        ref.append(REF_TEMPLATE.replace('__REF__', r))
ref = '\n'.join(ref)

csharp = re.sub(r'(namespace \S+)\n({)', r'''using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

// \1
// \2''', csharp)
csharp = re.sub(r'\n}\n', r'\n// }\n', csharp)
csharp = csharp.replace('public partial class MainForm : Form', f'public class {proj_name}Ribbon : ExcelRibbon')
csharp = csharp.replace('public MainForm()', 'public void RibbonLoad(IRibbonUI sender)')
csharp = csharp.replace('app = new XlApp();', 'app = (XlApp)ExcelDnaUtil.Application;')
csharp = csharp.replace('app.Visible = true;', '// app.Visible = true;')
csharp = csharp.replace('wb = (Workbook)app.Workbooks.Add();', 'wb = (Workbook)app.ActiveWorkbook;')
csharp = csharp.replace('ws = (Worksheet)wb.Worksheets.Add();', 'ws = (Worksheet)wb.ActiveSheet;')
csharp = csharp.replace('InitializeComponent();\n', '// InitializeComponent();\n')
csharp = re.sub(r'void (.*)\(object sender, EventArgs e\)', r'public void \1(IRibbonControl sender)', csharp)

for action in actions:
    if f'{action}(IRibbonControl ' not in csharp:
        print(f'*W: {action} is defined in xml, but not found in c#.')

# Output ------------------------------------------
if not os.path.exists(OUT_DIR):
    os.mkdir(OUT_DIR)
xll = OUT_DIR+'/'+proj_name+'.xll'
if not os.path.exists(xll):
    shutil.copyfile(XLL_PATH, xll)

dna = DNA_TEMPLATE.replace('__REFERENCE__', ref)
dna = dna.replace('__CSHARP__', csharp)
dna = dna.replace('__XML__', xml)
with open(OUT_DIR+'/'+proj_name+'.dna', 'w', encoding='utf8') as f:
    f.write(dna)
    
input('\n    Press ENTER to close ...')
