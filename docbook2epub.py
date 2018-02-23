#!/usr/bin/env python3

from __future__ import unicode_literals
from lxml import etree
import sys, os, os.path, logging, shutil, zipfile
import tkinter as tk
from tkinter.filedialog import askopenfilename
from threading import Thread

# not use yet, just for mark, for more, can see it at http://lxml.de/resolvers.html
class PrefixResolver(etree.Resolver):
    def __init__(self, prefix):
        self.prefix = prefix
        self.result_xml = '''\
             <xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
                 <test xmlns="testNS">%s-TEST</test>
             </xsl:stylesheet>
             ''' % prefix
    def resolve(self, url, pubid, context):
        if url.startswith(self.prefix):
            write_to_log("Resolved url %s as prefix %s" % (url, self.prefix))
            return self.resolve_string(self.result_xml, context)

class FileResolver(etree.Resolver):
    def resolve(self, url, pubid, context):
        return self.resolve_filename(url, context)

logging.basicConfig(level=logging.DEBUG)
log = logging.getLogger('docbook2epub')

DOCBOOK_XSL = os.path.abspath('./docbook-xsl-1.79.1/epub/docbook.xsl')
MIMETYPE = 'mimetype'
MIMETYPE_CONTENT = 'application/epub+zip'
TOOL_DIR = os.getcwd()

xslt_ac = etree.XSLTAccessControl(read_file=True,write_file=True, create_dir=True, read_network=True, write_network=False)
sourceforge_parser = etree.XMLParser()
sourceforge_parser.resolvers.add(FileResolver())
transform = etree.XSLT(etree.parse(DOCBOOK_XSL, sourceforge_parser), access_control=xslt_ac)

def async(f):
    def wrapper(*args, **kwargs):
        thr = Thread(target = f, args = args, kwargs = kwargs)
        thr.start()
    return wrapper

def fei(msg):
    print(msg)
    sys.exit()

def remove_biblioentry(docbook_file):
    tree = etree.ElementTree()
    tree.parse(docbook_file)
    root =  tree.getroot()
    for bibliography in root.findall("{http://docbook.org/ns/docbook}bibliography"):
        for biblioentry in bibliography.findall("{http://docbook.org/ns/docbook}biblioentry"):
            bibliography.remove(biblioentry)
    tree.write(docbook_file, encoding='UTF-8')
    return docbook_file

def convert_docbook(docbook_file):
    cwd = os.getcwd()
    docbook_dir = os.path.split(docbook_file)[0]
    output_path = os.path.join(docbook_dir, 'output')
    already_epub = os.path.join(docbook_dir, 'output.epub')
    if os.path.exists(already_epub):
        os.remove(already_epub)
    if os.path.exists(output_path):
        shutil.rmtree(output_path)
    if not os.path.exists(output_path):
        os.mkdir(output_path)
    shutil.copy(docbook_file, output_path)
    os.chdir(output_path)
    transform(etree.parse(docbook_file, sourceforge_parser), **{'html.stylesheet': "'css/stylesheets.css'"})

    os.chdir(cwd)
    return output_path

def find_resources(path):
    opf = etree.parse(os.path.join(path, 'OEBPS', 'content.opf'))
    for item in opf.xpath('//opf:item', namespaces= { 'opf': 'http://www.idpf.org/2007/opf' }):
        href = item.attrib['href']
        referenced_file = os.path.join(path, 'OEBPS', href)
        if not os.path.exists(referenced_file):
            log.debug("Copying '%s' into content folder" % href)
            #write_to_log("Copying '%s' into content folder\n" % href)
            try:
                dir = '%s/OEBPS/%s' % (path, os.path.split(href)[0].strip('./'))
                if not os.path.exists(dir):
                    write_to_log('mkdir: ' + dir)
                    os.mkdir(dir)
                shutil.copy(os.path.join(os.path.split(path)[0], href), dir)
            except FileNotFoundError as err:
                if 'stylesheets' in str(err):
                    shutil.copy(TOOL_DIR + '/stylesheets.css', dir)
                    write_to_log("[Tool]Copying '%s' into content folder" % href)
                else:
                    write_to_log(err)
    
def create_mimetype(path):
    f = '%s/%s' % (path, MIMETYPE)
    f = open(f, 'w')
    f.write(MIMETYPE_CONTENT)
    f.close()

def create_archive(path):
    output_dir = os.path.split(path)[0]

    epub_name = '%s.epub' % os.path.basename(path)
    os.chdir(path)    
    epub = zipfile.ZipFile(epub_name, 'w')
    epub.write(MIMETYPE, compress_type=zipfile.ZIP_STORED)

    for root, dirs, files in os.walk('.'):
        for file in files:
            if (not '.epub' in file and file != choose_file_name):
                epub.write(os.path.join(root, file), compress_type=zipfile.ZIP_DEFLATED)
    '''
    for p in os.listdir('.'):
        if os.path.isdir(p):
            for f in os.listdir(p):
                log.debug("Writing file '%s/%s'" % (p, f))
                #write_to_log("Writing file '%s/%s'\n" % (p, f))
                epub.write(os.path.join(p, f), compress_type=zipfile.ZIP_DEFLATED)
    '''
    epub.close()
    try:
        shutil.move(epub_name, output_dir)
    except shutil.Error as err:
        write_to_log(err)
    os.chdir(output_dir)
    
    return epub_name

@async
def convert(docbook_file):
    path = convert_docbook(docbook_file)
    find_resources(path)
    create_mimetype(path)
    epub = create_archive(path)
    #shutil.rmtree(path)

    write_to_log('[Output Dir:] ' + os.getcwd())
    os.chdir(TOOL_DIR)
    write_to_log("Created epub archive as '%s'\n" % epub)
    log.info("Created epub archive as '%s'" % epub)

def select_path():
    #default_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'docbook-to-epub')
    default_dir = os.path.join(os.path.expanduser('~'), 'Downloads', 'Main')
    main_xml = askopenfilename(initialdir = default_dir, title = 'Choose Main.xml', filetypes = (('Xml files', '*.xml'),('All files', '*.*')))
    path.set(main_xml)
    global choose_file_name
    choose_file_name = os.path.basename(main_xml)
    empty_log()
    write_to_log('[Choose Dir:] ' + os.path.split(main_xml)[0])
    convert(remove_biblioentry(main_xml))

def write_to_log(msg):
    numlines = log_print.index('end - 1 line').split('.')[0]
    if numlines == 24:
        log_print.delete(1.0, 2.0)
    if log_print.index('end-1c') != '1.0':
        log_print.insert('end', '\n')
    log_print.insert('end', msg)
    
def empty_log():
    log_print.delete(3.0, log_print.index('end - 2 line').split('.')[0] + '.0')

def git_shell(git_command):
    try:
      return os.popen(git_command).read().strip()
    except:
      return None

if __name__ == '__main__':
    git_shell('git pull')
    root = tk.Tk()
    root.title('Docbook Conversion Tool--MDPI')
    path = tk.StringVar()

    tk.Label(root,text = "Target Path:").grid(row = 0, column = 0, sticky = tk.W)
    tk.Entry(root, textvariable = path, width = 50).grid(row = 0, column = 1)
    tk.Button(root, text = "Choose Main.xml", command = select_path).grid(row = 0, column = 2)

    log_print = tk.Text(root, wrap = tk.WORD)
    log_print.grid(row = 1, columnspan= 3)
    write_to_log("============================================\n\t\t  Application start:\n============================================\n")

    root.mainloop()
