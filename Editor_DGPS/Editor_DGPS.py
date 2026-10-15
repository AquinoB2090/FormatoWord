import os, re, zipfile, tempfile, shutil, tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from xml.etree import ElementTree as ET

BASE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(BASE, 'Plantilla_DGPS.docx')
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
for prefix, uri in [('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main'),('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships'),('wp','http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'),('a','http://schemas.openxmlformats.org/drawingml/2006/main'),('pic','http://schemas.openxmlformats.org/drawingml/2006/picture'),('mc','http://schemas.openxmlformats.org/markup-compatibility/2006'),('w14','http://schemas.microsoft.com/office/word/2010/wordml'),('w15','http://schemas.microsoft.com/office/word/2012/wordml'),('wp14','http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'),('w10','urn:schemas-microsoft-com:office:word')]:
    ET.register_namespace(prefix, uri)

# Convierte cantidades comunes a palabras para mantener el estilo del formato.
UNIDADES = ['cero','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez','once','doce','trece','catorce','quince','dieciséis','diecisiete','dieciocho','diecinueve','veinte','veintiuno','veintidós','veintitrés','veinticuatro','veinticinco','veintiséis','veintisiete','veintiocho','veintinueve']
DECENAS = {30:'treinta',40:'cuarenta',50:'cincuenta',60:'sesenta',70:'setenta',80:'ochenta',90:'noventa'}
CENTENAS = {100:'cien',200:'doscientos',300:'trescientos',400:'cuatrocientos',500:'quinientos',600:'seiscientos',700:'setecientos',800:'ochocientos',900:'novecientos'}
def numero(n):
    if n < 30: return UNIDADES[n]
    if n < 100:
        d=(n//10)*10; return DECENAS[d] if n%10==0 else DECENAS[d]+' y '+UNIDADES[n%10]
    if n < 1000:
        c=(n//100)*100
        if n==100: return 'cien'
        pref='ciento' if c==100 else CENTENAS[c]
        return pref if n==c else pref+' '+numero(n-c)
    if n < 2000:
        return 'mil' if n==1000 else 'mil '+numero(n-1000)
    if n < 1000000:
        m=n//1000; rem=n%1000
        pref='mil' if m==1 else numero(m)+' mil'
        return pref if rem==0 else pref+' '+numero(rem)
    return str(n)
def fecha_en_palabras(value):
    d=datetime.strptime(value.strip(), '%d/%m/%Y')
    meses=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
    return f'{numero(d.day)} de {meses[d.month-1]} de {numero(d.year)}'

# lxml is shipped with standard Python on most Windows installations only in some setups;
# fall back to a small OOXML-safe implementation using ElementTree parent maps.
def replace_in_paragraph(p, old, new, global_replace=True):
    texts=list(p.iter(W+'t'))
    if not texts: return 0
    full=''.join(t.text or '' for t in texts)
    if not old: return 0
    pattern=re.compile(re.escape(old).replace(r'\ ', r'\s+'))
    matches=list(pattern.finditer(full))
    done=0
    for match in reversed(matches):
        starts=[]; pos=0
        for t in texts:
            s=t.text or ''; starts.append((pos,pos+len(s),t)); pos+=len(s)
        a,b=match.span()
        touched=[(s,e,t) for s,e,t in starts if e>a and s<b]
        if not touched: continue
        # These edits intentionally target the template's marked values. Accept only red runs.
        if not global_replace:
            parent={child:par for par in p.iter() for child in par}
            valid=True
            for _,_,t in touched:
                run=parent.get(t); rp=parent.get(run) if run is not None else None
                color=None
                if rp is not None:
                    color=next((c.attrib.get(W+'val','').upper() for c in rp if c.tag==W+'color'),None)
                if color not in ('EE0000','FF0000','RED'):
                    valid=False; break
            if not valid: continue
        s0,e0,t0=touched[0]; s1,e1,t1=touched[-1]
        local0=max(0,a-s0); local1=min(len(t1.text or ''),b-s1)
        if t0 is t1:
            value=t0.text or ''; t0.text=value[:local0]+new+value[local1:]
            if t0.text.startswith(' ') or t0.text.endswith(' '): t0.set(XML_SPACE,'preserve')
        else:
            v0=t0.text or ''; t0.text=v0[:local0]+new
            for _,_,t in touched[1:-1]: t.text=''
            v1=t1.text or ''; t1.text=v1[local1:]
            if t0.text.startswith(' ') or t0.text.endswith(' '): t0.set(XML_SPACE,'preserve')
            if (t1.text or '').startswith(' ') or (t1.text or '').endswith(' '): t1.set(XML_SPACE,'preserve')
        # refresh text list and flattened content before applying earlier matches
        texts=list(p.iter(W+'t'))
        full=''.join(t.text or '' for t in texts)
        done+=1
    return done

def edit_docx(template, output, values):
    with zipfile.ZipFile(template,'r') as src:
        entries=[(i,src.read(i.filename)) for i in src.infolist()]
    replacements=[
        ('Justiniano Carrera',values['nombre']),
        ('ochenta y cinco años',numero(int(values['edad']))+' años'),
        ('Casado',values['estado']),
        ('Aldea Suquinay',values['direccion']),
        ('San Antonio La Paz',values['municipio']),
        ('El Progreso',values['departamento']),
        ('2716102710207',values['dpi']),
        ('42073950',values['telefono']),
        ('DGPS-B-523591219520082024',values['numero']),
        ('trece de enero  de dos mil veintiséis',fecha_en_palabras(values['fecha_resolucion'])),
        ('trece de enero  de dos mil veintiséis',fecha_en_palabras(values['fecha_resolucion'])),
        ('día quince de  octubre del año dos mil veintiséis',fecha_en_palabras(values['fecha_notificacion'])),
        ('quince de  octubre del año dos mil veintiséis',fecha_en_palabras(values['fecha_notificacion'])),
        ('día quince de octubre del año dos mil veintiséis',fecha_en_palabras(values['fecha_notificacion'])),
        ('En el año dos mil veinticuatro','En el año '+numero(int(values['anio_solicitud']))),
    ]
    # More specific replacements precede the year-only substitution.
    replacements.sort(key=lambda item:len(item[0]), reverse=True)
    outdir=os.path.dirname(os.path.abspath(output)) or BASE
    fd,tmp=tempfile.mkstemp(suffix='.docx',dir=outdir); os.close(fd)
    try:
        with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as dst:
            for info,data in entries:
                if re.fullmatch(r'word/(document|header\d*|footer\d*)\.xml',info.filename):
                    try:
                        root=ET.fromstring(data)
                        for p in root.iter(W+'p'):
                            for old,new in replacements:
                                replace_in_paragraph(p,old,new,global_replace=True)
                        data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
                    except ET.ParseError:
                        pass
                dst.writestr(info,data)
        shutil.move(tmp,output)
    finally:
        if os.path.exists(tmp): os.remove(tmp)

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title('Editor de formato DGPS'); self.geometry('690x700'); self.minsize(620,620)
        self.configure(bg='#f4f6f8')
        style=ttk.Style(self); style.theme_use('clam'); style.configure('TLabel',background='#f4f6f8',font=('Segoe UI',10)); style.configure('Title.TLabel',font=('Segoe UI',17,'bold'),foreground='#17365d'); style.configure('TButton',font=('Segoe UI',10),padding=8)
        top=ttk.Frame(self,padding=(22,18,22,8)); top.pack(fill='x'); ttk.Label(top,text='Completar formato DGPS',style='Title.TLabel').pack(anchor='w'); ttk.Label(top,text='Escribe los datos de la persona y guarda un Word nuevo.').pack(anchor='w',pady=(4,0))
        shell=ttk.Frame(self,padding=(22,8,22,10)); shell.pack(fill='both',expand=True)
        canvas=tk.Canvas(shell,highlightthickness=0,bg='#f4f6f8'); bar=ttk.Scrollbar(shell,orient='vertical',command=canvas.yview); self.form=ttk.Frame(canvas); self.form.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all'))); canvas.create_window((0,0),window=self.form,anchor='nw'); canvas.configure(yscrollcommand=bar.set); canvas.pack(side='left',fill='both',expand=True); bar.pack(side='right',fill='y')
        self.fields=[('nombre','Nombre completo','Justiniano Carrera'),('edad','Edad (años)','85'),('estado','Estado civil','Casado'),('dpi','DPI / CUI','2716102710207'),('direccion','Dirección de residencia','Aldea Suquinay'),('municipio','Municipio','San Antonio La Paz'),('departamento','Departamento','El Progreso'),('telefono','Teléfono','42073950'),('numero','Número de resolución/formulario','DGPS-B-523591219520082024'),('fecha_resolucion','Fecha de resolución (dd/mm/aaaa)','13/01/2026'),('fecha_notificacion','Fecha de notificación (dd/mm/aaaa)','15/10/2026'),('anio_solicitud','Año de solicitud','2024')]
        self.vars={}
        for i,(key,label,default) in enumerate(self.fields):
            ttk.Label(self.form,text=label).grid(row=i,column=0,sticky='w',pady=5,padx=(0,16)); v=tk.StringVar(value=default); self.vars[key]=v; ttk.Entry(self.form,textvariable=v,width=46).grid(row=i,column=1,sticky='ew',pady=5)
        self.form.columnconfigure(1,weight=1)
        foot=ttk.Frame(self,padding=(22,8,22,20)); foot.pack(fill='x'); ttk.Button(foot,text='Crear documento Word…',command=self.create).pack(side='right')
    def create(self):
        vals={k:v.get().strip() for k,v in self.vars.items()}
        try:
            if not vals['nombre'] or not vals['estado'] or not vals['numero'] or not vals['direccion']: raise ValueError('Completa al menos nombre, estado civil, dirección y número de resolución.')
            for k in ('edad','anio_solicitud'): int(vals[k])
            for k in ('fecha_resolucion','fecha_notificacion'): datetime.strptime(vals[k],'%d/%m/%Y')
            if not os.path.isfile(TEMPLATE): raise ValueError('No encuentro Plantilla_DGPS.docx junto al programa.')
        except ValueError as e:
            messagebox.showerror('Revisa los datos',str(e)); return
        path=filedialog.asksaveasfilename(title='Guardar documento completado',defaultextension='.docx',initialfile='Recurso_DGPS_'+vals['nombre'].replace(' ','_')+'.docx',filetypes=[('Documento Word','*.docx')])
        if not path: return
        try:
            edit_docx(TEMPLATE,path,vals); messagebox.showinfo('Documento creado','Se guardó el documento en:\n'+path)
        except Exception as e: messagebox.showerror('No se pudo crear',str(e))

if __name__=='__main__':
    App().mainloop()

