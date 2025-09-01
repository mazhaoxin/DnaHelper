# MaZhaoxin Re-write the code @20250628.
# Features:
#   v0.1: Convert xxx.dna from xxx/MainForm.cs.
#   v0.2: Support standalone C# file.
#   v0.3: Read refernce from xxx.csproj.

# MaZhaoxin @20230507
# Usage:
#   1. Copy this script to *.sln located level.
#   2. Run it.
#   3. Find *.xll, *.dna and auto_load.vbs in `./Distribution`.
# Note:
#   1. Must involve Excel namespace and rename `Excel.Application` to `ExcelApp`.
#   2. Must delare `app`, `wb` and `ws` for `Excel.Application`, `Workbook` and `Worksheet`.
#   3. Must init them in constructor by
#       app = new ExcelApp();
#       app.Visible = true;
#       wb = (Workbook)app.Workbooks.Add();
#       ws = (Worksheet)wb.Worksheets.Add();
# Warning:
#   1. Cannot use $"{x}" as that it only can work after C# 6.0.

XLL_PATH = r'E:\Programming\ExcelDNA\ExcelDna64.xll'
OUT_DIR = 'Distribution'


import os, shutil, sys
import re
import logging
logging.basicConfig(level=logging.DEBUG)


def is_exist(name:str) -> bool:
    return os.path.exists(name)

def get_proj_name() -> str:
    cwd = os.path.abspath(os.curdir)
    proj_name = cwd.split('\\')[-1]
    return proj_name

def is_csproj(proj_name:str) -> bool:
    return is_exist(f'./{proj_name}/{proj_name}.csproj')

def get_ref(proj_name:str) -> list[str]:
    if not is_csproj(proj_name):
        return [
            'System.Windows.Forms',
            'Microsoft.Office.Interop.Excel',
        ]
    
    from xml.dom import minidom

    # parse xxx.csproj and find all Reference, then get Include attribute.
    doc = minidom.parse(f'./{proj_name}/{proj_name}.csproj')
    ref_list = []
    sys_list = [
        'Microsoft.CSharp',
        'System', 
        'System.Core', 
        'System.Data', 
        'System.Data.DataSetExtensions', 
        'System.Drawing', 
        'System.Xml', 
        'System.Xml.Linq',
    ]
    for ref in doc.getElementsByTagName('Reference'):
        if ref.hasAttribute('Include'):
            ref_inc = ref.getAttribute('Include')
            if ref_inc not in sys_list:
                ref_list.append(ref_inc)
    return ref_list

def get_cs(proj_name:str) -> str:
    if is_csproj(proj_name):
        fpath = f'./{proj_name}/MainForm.cs'
    else:
        fpath = 'MainForm.cs'
    if not is_exist(fpath):
        write_cs_template(proj_name)

    cs = open(fpath, encoding='utf8').read()
    return cs

def get_cs_actions(cs:str) -> list[str]:
    pattern = r'\bvoid\s+([A-Za-z_]\w*)\s*\(\s*object\s+sender\s*,\s*EventArgs\s+e\s*\)'
    actions = []
    matches = re.findall(pattern, cs)
    actions.extend(matches)
    return actions

def get_xml(proj_name:str) -> str:
    if not is_exist(proj_name+'.xml'):
        write_xml_template(proj_name)
    xml = open(proj_name+'.xml', encoding='utf8').read()
    return xml

def get_xml_actions(xml:str) -> list[str]:
    from collections import namedtuple
    ActionInfo = namedtuple('ActionInfo', ['Type', 'ID', 'Action', 'Callback'])

    from xml.dom import minidom
    doc = minidom.parseString(xml)
    
    def traverse(element:minidom.Element) -> list[ActionInfo]:
        results = []
        # check all attributes
        if element.attributes:
            for i in range(element.attributes.length):
                attr = element.attributes.item(i)
                if attr.name.startswith('on'):
                    results.append(ActionInfo(
                        Type=element.tagName,
                        ID=element.getAttribute('id'),
                        Action=attr.name,
                        Callback=attr.value
                    ))
        # check all children
        for child in element.childNodes:
            if child.nodeType == child.ELEMENT_NODE:
                results.extend(traverse(child))
        return results
    
    infos = traverse(doc.documentElement)
    logging.debug('---- Actions from XML ----')
    for info in infos:
        logging.debug(info)
    logging.debug('---- ---------------- ----')
    return [x.Callback for x in infos]

def convert_ref(ref:list[str]) -> str:
    REF_TEMPLATE = '<Reference Name="__REF__" />'
    rxml = []
    for r in ref:
        rxml.append(REF_TEMPLATE.replace('__REF__', r))
    return '\n'.join(rxml)

def convert_cs(cs:str) -> str:
    # insert ExcelDna namespace
    cs = re.sub(r'(namespace \S+)\n({)', r'''using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

// \1
// \2''', cs)
    cs = re.sub(r'\n}\n', r'\n// }\n', cs)
    cs = cs.replace('public partial class MainForm : Form', f'public class {proj_name}Ribbon : ExcelRibbon')
    cs = cs.replace('public MainForm()', 'public void RibbonLoad(IRibbonUI sender)')
    cs = cs.replace('app = new ExcelApp();', 'app = (ExcelApp)ExcelDnaUtil.Application;')
    cs = cs.replace('app.Visible = true;', '// app.Visible = true;')
    cs = cs.replace('wb = (Workbook)app.Workbooks.Add();', 'wb = (Workbook)app.ActiveWorkbook;')
    cs = cs.replace('ws = (Worksheet)wb.Worksheets.Add();', 'ws = (Worksheet)wb.ActiveSheet;')
    cs = cs.replace('InitializeComponent();\n', '// InitializeComponent();\n')
    cs = re.sub(r'void (.*)\(object sender, EventArgs e\)', r'public void \1(IRibbonControl sender)', cs)
    return cs

def write_dna(proj_name:str, ref:str, cs:str, xml:str) -> None:
    if not os.path.exists(OUT_DIR):
        os.mkdir(OUT_DIR)
    xll_path = OUT_DIR+'/'+proj_name+'.xll'
    dna_path = OUT_DIR+'/'+proj_name+'.dna'

    if not os.path.exists(xll_path):
        shutil.copyfile(XLL_PATH, xll_path)

    DNA_TEMPLATE = '''<DnaLibrary Name="__NAME__" RuntimeVersion="v4.0" Language="C#" >
__REFERENCE__
<![CDATA[
__CSHARP__
]]>
<CustomUI>
__XML__
</CustomUI>
</DnaLibrary>'''
    dna = DNA_TEMPLATE.replace('__NAME__', proj_name)
    dna = dna.replace('__REFERENCE__', ref)
    dna = dna.replace('__CSHARP__', cs)
    dna = dna.replace('__XML__', xml)
    with open(dna_path, 'w', encoding='utf8') as f:
        f.write(dna)

def write_vbs() -> None:
    VBS = '''Application.RegisterXLL...
'''
    pass

def write_cs_template(proj_name:str) -> None:
    CS_TEMPLATE = '''using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace SimpleExample
{
	public partial class MainForm : Form
	{
		ExcelApp app;
		Workbook wb;
		Worksheet ws;
		
		public MainForm()
		{
			InitializeComponent();
			
			app = new ExcelApp();
			app.Visible = true;
			wb = (Workbook)app.Workbooks.Add();
			ws = (Worksheet)wb.Worksheets.Add();
		}
		
		void Button1Click(object sender, EventArgs e)
		{
			// TODO: Implement Button1Click
			MessageBox.Show("Button-1", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
		
		void Button2Click(object sender, EventArgs e)
		{
			// TODO: Implement Button2Click
			MessageBox.Show("Button-2", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}
}
'''
    if not is_exist('MainForm.cs'):
        with open('MainForm.cs', 'w', encoding='utf8') as f:
            f.write(CS_TEMPLATE)

def write_xml_template(proj_name:str) -> None:
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
    if not is_exist(proj_name+'.xml'):
        with open(proj_name+'.xml', 'w', encoding='utf8') as f:
            f.write(XML_TEMPLATE)

def check_actions(xml_actions:list[str], cs_actions:list[str]) -> None:
    if 'RibbonLoad' in xml_actions:
        xml_actions.remove('RibbonLoad')
    xml_actions = set(xml_actions)
    cs_actions = set(cs_actions)
    
    if xml_actions==cs_actions:
        logging.info('Actions compare pass.')
        return
    
    for a in xml_actions:
        if a not in cs_actions:
            logging.warning(f'{a} is used in XML, but not declared in C#.')
    for a in cs_actions:
        if a not in xml_actions:
            logging.info(f'{a} is declared in C#, but not used in XML.')

if __name__ == '__main__':
    proj_name = get_proj_name()
    logging.info(f'Project Name: {proj_name}')

    ref = get_ref(proj_name)

    xml = get_xml(proj_name)
    xml_actions = get_xml_actions(xml)
    logging.debug(f'Actions from XML: {xml_actions}')

    cs = get_cs(proj_name)
    cs_actions = get_cs_actions(cs)
    logging.debug(f'Actions from C#: {cs_actions}')

    check_actions(xml_actions, cs_actions)

    write_dna(
        proj_name,
        convert_ref(ref),
        convert_cs(cs),
        xml
    )

    input('\n    Press ENTER to close ...')

