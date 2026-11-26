"""PRIMARIA NEZAHUALCOYOTL | CONTROL ESCOLAR
Aplicación de un solo archivo: interfaz Tkinter, esquema MySQL y ayuda de instalación.
Requisitos externos: Windows 10/11, Python 3.11+, XAMPP/MySQL.
Instalación: py -m pip install mysql-connector-python openpyxl
Ejecución: py app.py
MySQL se configura con NEZA_DB_HOST, NEZA_DB_PORT, NEZA_DB_USER, NEZA_DB_PASSWORD.
Cuentas de demostración: director/director y docente1/docente1 ... docente6/docente6.
Cambie todas las contraseñas de demostración antes de usar información real.
Excel: encabezados requeridos nombre, grupo; opcionales numero_lista, tutor, telefono, correo.
La aplicación crea la base de datos y las tablas automáticamente.
"""
from __future__ import annotations
import os, sys, hashlib, hmac, secrets, tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from datetime import date, datetime

APP_TITLE = 'PRIMARIA NEZAHUALCÓYOTL | CONTROL ESCOLAR'
DB_NAME = os.getenv('NEZA_DB_NAME', 'control_escolar_neza')
DB_CONFIG = dict(host=os.getenv('NEZA_DB_HOST','127.0.0.1'), port=int(os.getenv('NEZA_DB_PORT','3306')),
                 user=os.getenv('NEZA_DB_USER','root'), password=os.getenv('NEZA_DB_PASSWORD',''),
                 charset='utf8mb4', use_unicode=True)
# Cada elemento visual tiene su propia clave para personalización independiente.
THEME = {
 'login_blue':'#19324D','login_white':'#FFFFFF','login_accent':'#77C9D0','background':'#F3F6F8',
 'header':'#19324D','sidebar':'#19324D','sidebar_text':'#FFFFFF','sidebar_hover':'#28516F',
 'title':'#19324D','subtitle':'#64748B','card_present_bg':'#E7F5EC','card_present_fg':'#207547',
 'card_absence_bg':'#FBE9E9','card_absence_fg':'#B42318','card_late_bg':'#FFF3D9','card_late_fg':'#946200',
 'card_neutral_bg':'#FFFFFF','card_neutral_fg':'#19324D','button_primary_bg':'#167D8D',
 'button_primary_fg':'#FFFFFF','button_secondary_bg':'#E5F2F8','button_secondary_fg':'#19324D',
 'table_bg':'#FFFFFF','table_heading_bg':'#D7E8F0','table_fg':'#19324D','table_selected':'#B9E5EA',
 'present_row':'#E7F5EC','absence_row':'#FBE9E9','late_row':'#FFF3D9','border':'#DCE5EA',
 'panel':'#FFFFFF','success':'#207547','danger':'#B42318','warning':'#946200'
}

def hash_pw(password:str)->str:
    salt=secrets.token_hex(16)
    digest=hashlib.pbkdf2_hmac('sha256',password.encode(),salt.encode(),240_000).hex()
    return f'pbkdf2_sha256${salt}${digest}'

def verify_pw(password:str, stored:str)->bool:
    """Verifica hashes actuales y hashes de la versión anterior (200,000 iteraciones)."""
    try:
        scheme,salt,digest=stored.split('$',2)
        if scheme!='pbkdf2_sha256':
            return False
        for iterations in (240_000, 200_000):
            test=hashlib.pbkdf2_hmac('sha256',password.encode(),salt.encode(),iterations).hex()
            if hmac.compare_digest(test,digest):
                return True
        return False
    except Exception:
        return False

def connect(database=True):
    try:
        import mysql.connector
        cfg=dict(DB_CONFIG)
        if database: cfg['database']=DB_NAME
        return mysql.connector.connect(**cfg)
    except Exception as e:
        raise RuntimeError('No se pudo conectar a MySQL. Enciende MySQL en XAMPP y revisa la configuración NEZA_DB_*.\n\n'+str(e))

def query(sql, params=()):
    con=connect()
    try:
        cur=con.cursor(dictionary=True); cur.execute(sql,params); return cur.fetchall()
    finally: con.close()

def execute(sql, params=()):
    con=connect()
    try:
        cur=con.cursor(); cur.execute(sql,params); con.commit(); return cur.lastrowid
    except Exception: con.rollback(); raise
    finally: con.close()

def executemany(sql, rows):
    con=connect()
    try:
        cur=con.cursor(); cur.executemany(sql,rows); con.commit(); return cur.rowcount
    except Exception: con.rollback(); raise
    finally: con.close()

DDL = [
"""CREATE TABLE IF NOT EXISTS usuarios(
 id INT AUTO_INCREMENT PRIMARY KEY, usuario VARCHAR(50) NOT NULL UNIQUE,
 password_hash VARCHAR(255) NOT NULL, rol ENUM('director','docente','tutor') NOT NULL,
 nombre VARCHAR(150) NOT NULL, telefono VARCHAR(30) NULL, correo VARCHAR(120) NULL,
 direccion VARCHAR(220) NULL, activo TINYINT(1) NOT NULL DEFAULT 1,
 creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP) ENGINE=InnoDB""",
"""CREATE TABLE IF NOT EXISTS grupos(
 id INT AUTO_INCREMENT PRIMARY KEY, nombre VARCHAR(40) NOT NULL UNIQUE,
 docente_id INT NULL, activo TINYINT(1) NOT NULL DEFAULT 1,
 CONSTRAINT fk_grupo_docente FOREIGN KEY(docente_id) REFERENCES usuarios(id) ON DELETE SET NULL) ENGINE=InnoDB""",
"""CREATE TABLE IF NOT EXISTS alumnos(
 id INT AUTO_INCREMENT PRIMARY KEY, grupo_id INT NOT NULL, numero_lista INT NOT NULL DEFAULT 0,
 nombre VARCHAR(150) NOT NULL, matricula VARCHAR(40) NULL, fecha_nacimiento DATE NULL,
 tutor_nombre VARCHAR(150) NULL, tutor_telefono VARCHAR(30) NULL, tutor_correo VARCHAR(120) NULL,
 tutor_usuario_id INT NULL, domicilio VARCHAR(220) NULL, observaciones TEXT NULL,
 activo TINYINT(1) NOT NULL DEFAULT 1, UNIQUE KEY uq_lista(grupo_id,numero_lista),
 KEY idx_alumno_nombre(nombre), CONSTRAINT fk_alumno_grupo FOREIGN KEY(grupo_id) REFERENCES grupos(id) ON DELETE RESTRICT,
 CONSTRAINT fk_alumno_tutor FOREIGN KEY(tutor_usuario_id) REFERENCES usuarios(id) ON DELETE SET NULL) ENGINE=InnoDB""",
"""CREATE TABLE IF NOT EXISTS asistencias(
 id INT AUTO_INCREMENT PRIMARY KEY, alumno_id INT NOT NULL, grupo_id INT NOT NULL, docente_id INT NOT NULL,
 fecha DATE NOT NULL, estado ENUM('Presente','Falta','Retardo') NOT NULL DEFAULT 'Presente',
 actualizado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
 UNIQUE KEY uq_asistencia(alumno_id,fecha), KEY idx_asistencia_fecha(fecha,grupo_id),
 FOREIGN KEY(alumno_id) REFERENCES alumnos(id) ON DELETE RESTRICT,
 FOREIGN KEY(grupo_id) REFERENCES grupos(id) ON DELETE RESTRICT,
 FOREIGN KEY(docente_id) REFERENCES usuarios(id) ON DELETE RESTRICT) ENGINE=InnoDB""",
"""CREATE TABLE IF NOT EXISTS justificantes(
 id INT AUTO_INCREMENT PRIMARY KEY, asistencia_id INT NOT NULL UNIQUE, motivo VARCHAR(600) NOT NULL,
 estado ENUM('Pendiente','Aprobado','Rechazado') NOT NULL DEFAULT 'Pendiente', respuesta TEXT NULL,
 creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP, revisado_en DATETIME NULL,
 FOREIGN KEY(asistencia_id) REFERENCES asistencias(id) ON DELETE CASCADE) ENGINE=InnoDB""",
"""CREATE TABLE IF NOT EXISTS incidencias(
 id INT AUTO_INCREMENT PRIMARY KEY, alumno_id INT NOT NULL, grupo_id INT NOT NULL, docente_id INT NOT NULL,
 tipo ENUM('Conducta','Incidencia') NOT NULL DEFAULT 'Incidencia', titulo VARCHAR(160) NOT NULL,
 gravedad ENUM('Leve','Media','Grave') NOT NULL DEFAULT 'Leve', descripcion TEXT NOT NULL,
 medida TEXT NULL, fecha DATE NOT NULL, creada_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
 FOREIGN KEY(alumno_id) REFERENCES alumnos(id) ON DELETE RESTRICT,
 FOREIGN KEY(grupo_id) REFERENCES grupos(id) ON DELETE RESTRICT,
 FOREIGN KEY(docente_id) REFERENCES usuarios(id) ON DELETE RESTRICT) ENGINE=InnoDB""",
"""CREATE TABLE IF NOT EXISTS mensajes(
 id INT AUTO_INCREMENT PRIMARY KEY, remitente_id INT NOT NULL, destinatario_id INT NOT NULL,
 tipo ENUM('DIRECTOR_DOCENTE','DOCENTE_DIRECTOR') NOT NULL, titulo VARCHAR(180) NOT NULL,
 mensaje TEXT NOT NULL, leido TINYINT(1) NOT NULL DEFAULT 0, creado_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
 KEY idx_mensaje_destino(destinatario_id,leido,creado_en),
 FOREIGN KEY(remitente_id) REFERENCES usuarios(id) ON DELETE CASCADE,
 FOREIGN KEY(destinatario_id) REFERENCES usuarios(id) ON DELETE CASCADE) ENGINE=InnoDB""",
"""CREATE TABLE IF NOT EXISTS cartas_compromiso(
 id INT AUTO_INCREMENT PRIMARY KEY, alumno_id INT NOT NULL, grupo_id INT NOT NULL, director_id INT NOT NULL,
 incidencia_id INT NULL, fecha DATE NOT NULL, titulo VARCHAR(150) NOT NULL, motivo TEXT NOT NULL,
 compromisos TEXT NOT NULL, fecha_seguimiento DATE NULL,
 estado ENUM('Pendiente','En seguimiento','Cumplida','Cerrada') NOT NULL DEFAULT 'Pendiente',
 creada_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
 FOREIGN KEY(alumno_id) REFERENCES alumnos(id) ON DELETE RESTRICT,
 FOREIGN KEY(grupo_id) REFERENCES grupos(id) ON DELETE RESTRICT,
 FOREIGN KEY(director_id) REFERENCES usuarios(id) ON DELETE RESTRICT,
 FOREIGN KEY(incidencia_id) REFERENCES incidencias(id) ON DELETE SET NULL) ENGINE=InnoDB""",
"""CREATE TABLE IF NOT EXISTS notificaciones(
 id INT AUTO_INCREMENT PRIMARY KEY, destinatario_id INT NOT NULL, emisor_id INT NULL,
 tipo VARCHAR(50) NOT NULL, titulo VARCHAR(180) NOT NULL, mensaje TEXT NOT NULL,
 referencia_id INT NULL, prioridad ENUM('normal','alta','urgente') NOT NULL DEFAULT 'normal',
 leida TINYINT(1) NOT NULL DEFAULT 0, creada_en TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
 KEY idx_noti_destino(destinatario_id,leida,creada_en),
 FOREIGN KEY(destinatario_id) REFERENCES usuarios(id) ON DELETE CASCADE,
 FOREIGN KEY(emisor_id) REFERENCES usuarios(id) ON DELETE SET NULL) ENGINE=InnoDB"""
]

def initialize_db():
    con=connect(False)
    try:
        cur=con.cursor(); cur.execute(f"CREATE DATABASE IF NOT EXISTS `{DB_NAME}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"); con.commit()
    finally: con.close()
    con=connect()
    try:
        cur=con.cursor()
        for statement in DDL: cur.execute(statement)
        con.commit()
        # Migraciones pequeñas para bases creadas con versiones anteriores.
        for table, col, definition in [
          ('usuarios','telefono','VARCHAR(30) NULL'),('usuarios','correo','VARCHAR(120) NULL'),('usuarios','direccion','VARCHAR(220) NULL'),
          ('alumnos','matricula','VARCHAR(40) NULL'),('alumnos','fecha_nacimiento','DATE NULL'),('alumnos','tutor_nombre','VARCHAR(150) NULL'),
          ('alumnos','tutor_telefono','VARCHAR(30) NULL'),('alumnos','tutor_correo','VARCHAR(120) NULL'),('alumnos','domicilio','VARCHAR(220) NULL'),('alumnos','observaciones','TEXT NULL'),
          ('justificantes','respuesta','TEXT NULL'),('justificantes','revisado_en','DATETIME NULL'),
          ('notificaciones','emisor_id','INT NULL')]:
            cur.execute("SELECT COUNT(*) FROM information_schema.COLUMNS WHERE TABLE_SCHEMA=%s AND TABLE_NAME=%s AND COLUMN_NAME=%s",(DB_NAME,table,col))
            if cur.fetchone()[0]==0: cur.execute(f"ALTER TABLE `{table}` ADD COLUMN `{col}` {definition}")
        con.commit()
    finally: con.close()
    accounts=[('director','director','director','Dirección escolar')]+[(f'docente{i}',f'docente{i}','docente',f'Docente {i}') for i in range(1,7)]
    for u,pw,role,name in accounts:
        if not query('SELECT id FROM usuarios WHERE usuario=%s',(u,)):
            execute('INSERT INTO usuarios(usuario,password_hash,rol,nombre) VALUES(%s,%s,%s,%s)',(u,hash_pw(pw),role,name))
    teachers=query("SELECT id,usuario FROM usuarios WHERE rol='docente' AND activo=1")
    for i in range(1,7):
        name=f'Grupo {i}'; g=query('SELECT id FROM grupos WHERE nombre=%s',(name,))
        teacher=next((x['id'] for x in teachers if x['usuario']==f'docente{i}'),None)
        if not g: execute('INSERT INTO grupos(nombre,docente_id) VALUES(%s,%s)',(name,teacher))
        elif teacher and not query('SELECT id FROM grupos WHERE docente_id=%s',(teacher,)): execute('UPDATE grupos SET docente_id=%s WHERE id=%s',(teacher,g[0]['id']))

def today(): return date.today().isoformat()
def fmtdate(v): return v.strftime('%d/%m/%Y') if hasattr(v,'strftime') else str(v or '')

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP_TITLE); self.geometry('1360x850'); self.minsize(1050,680)
        try:self.state('zoomed')
        except tk.TclError:pass
        self.configure(bg=THEME['background']); self.user=None; self.content=None; self.setup_style(); self.login()
    def setup_style(self):
        s=ttk.Style(self); s.theme_use('clam')
        s.configure('Treeview',background=THEME['table_bg'],fieldbackground=THEME['table_bg'],foreground=THEME['table_fg'],rowheight=29,font=('Segoe UI',10))
        s.configure('Treeview.Heading',background=THEME['table_heading_bg'],foreground=THEME['table_fg'],font=('Segoe UI',10,'bold'),padding=7)
        s.map('Treeview',background=[('selected',THEME['table_selected'])],foreground=[('selected',THEME['table_fg'])])
        s.configure('TCombobox',padding=6)
    def clear(self):
        for w in self.winfo_children(): w.destroy()
    def btn(self,parent,text,cmd,primary=False):
        return tk.Button(parent,text=text,command=cmd,bg=THEME['button_primary_bg'] if primary else THEME['button_secondary_bg'],fg=THEME['button_primary_fg'] if primary else THEME['button_secondary_fg'],activebackground=THEME['sidebar_hover'],activeforeground='white',relief='flat',bd=0,cursor='hand2',font=('Segoe UI',9,'bold'),padx=12,pady=8)
    def login(self):
        self.user=None; self.clear(); base=tk.Frame(self,bg=THEME['login_white']); base.pack(fill='both',expand=True)
        left=tk.Frame(base,bg=THEME['login_blue']); left.place(relx=0,rely=0,relwidth=.5,relheight=1)
        # Decoración escolar dentro de la mitad azul.
        tk.Canvas(left,bg=THEME['login_blue'],highlightthickness=0).place(relx=0,rely=0,relwidth=1,relheight=1)
        tk.Label(left,text='✦',bg=THEME['login_blue'],fg=THEME['login_accent'],font=('Segoe UI',52)).place(relx=.14,rely=.13)
        tk.Label(left,text='NZ',bg=THEME['login_blue'],fg=THEME['login_accent'],font=('Segoe UI',62,'bold')).place(relx=.5,rely=.28,anchor='center')
        tk.Label(left,text='PRIMARIA\nNEZAHUALCÓYOTL',bg=THEME['login_blue'],fg='white',font=('Segoe UI',27,'bold'),justify='center').place(relx=.5,rely=.49,anchor='center')
        tk.Label(left,text='CONTROL ESCOLAR',bg=THEME['login_blue'],fg='#BFE5EA',font=('Segoe UI',15,'bold')).place(relx=.5,rely=.64,anchor='center')
        tk.Label(left,text='Asistencia  •  Seguimiento  •  Gestión',bg=THEME['login_blue'],fg='#D5E5F1',font=('Segoe UI',11)).place(relx=.5,rely=.70,anchor='center')
        tk.Label(left,text='✧                                      ✧',bg=THEME['login_blue'],fg='#47718B',font=('Segoe UI',28)).place(relx=.5,rely=.86,anchor='center')
        right=tk.Frame(base,bg=THEME['login_white']); right.place(relx=.5,rely=0,relwidth=.5,relheight=1)
        form=tk.Frame(right,bg='white'); form.place(relx=.5,rely=.5,anchor='center',relwidth=.68)
        tk.Label(form,text='Bienvenido',bg='white',fg=THEME['title'],font=('Segoe UI',29,'bold')).pack(anchor='w',pady=(0,6))
        tk.Label(form,text='Inicia sesión para ingresar al sistema escolar',bg='white',fg=THEME['subtitle'],font=('Segoe UI',11)).pack(anchor='w',pady=(0,28))
        tk.Label(form,text='Usuario',bg='white',fg=THEME['title'],font=('Segoe UI',10,'bold')).pack(anchor='w')
        self.user_entry=tk.Entry(form,font=('Segoe UI',12),relief='solid',bd=1); self.user_entry.pack(fill='x',ipady=9,pady=(5,18))
        tk.Label(form,text='Contraseña',bg='white',fg=THEME['title'],font=('Segoe UI',10,'bold')).pack(anchor='w')
        self.pass_entry=tk.Entry(form,font=('Segoe UI',12),show='●',relief='solid',bd=1); self.pass_entry.pack(fill='x',ipady=9,pady=(5,22))
        self.btn(form,'Iniciar sesión',self.do_login,True).pack(fill='x',ipady=3)
        tk.Label(form,text='Acceso autorizado para personal y tutores',bg='white',fg='#94A3B8',font=('Segoe UI',9)).pack(anchor='center',pady=22)
        self.pass_entry.bind('<Return>',lambda e:self.do_login()); self.user_entry.focus_set()
    def do_login(self):
        u=self.user_entry.get().strip(); pw=self.pass_entry.get()
        if not u or not pw: messagebox.showwarning('Datos','Escribe usuario y contraseña.'); return
        try:
            rows=query('SELECT * FROM usuarios WHERE usuario=%s AND activo=1',(u,))
            if not rows or not verify_pw(pw,rows[0]['password_hash']): messagebox.showerror('Acceso','Usuario o contraseña incorrectos.'); return
            self.user=rows[0]; self.dashboard()
        except Exception as e: messagebox.showerror('Conexión',str(e))
    def dashboard(self):
        self.clear(); top=tk.Frame(self,bg=THEME['header'],height=62); top.pack(fill='x'); top.pack_propagate(False)
        tk.Label(top,text=APP_TITLE,bg=THEME['header'],fg='white',font=('Segoe UI',12,'bold')).pack(side='left',padx=22)
        tk.Label(top,text=f"{self.user['nombre']}  |  {self.user['rol'].upper()}",bg=THEME['header'],fg='#DCE8EF',font=('Segoe UI',10)).pack(side='right',padx=15)
        self.btn(top,'Cerrar sesión',self.login).pack(side='right',padx=18,pady=12)
        shell=tk.Frame(self,bg=THEME['background']); shell.pack(fill='both',expand=True,padx=14,pady=14)
        side=tk.Frame(shell,bg=THEME['sidebar'],width=225); side.pack(side='left',fill='y',padx=(0,14)); side.pack_propagate(False)
        self.content=tk.Frame(shell,bg=THEME['background']); self.content.pack(side='left',fill='both',expand=True)
        role=self.user['rol']
        if role=='director': items=[('Resumen y asistencia',self.summary),('Alumnos',self.students),('Docentes y grupos',self.groups),('Conductas recibidas',lambda:self.incidents_admin('Conducta')),('Incidencias recibidas',lambda:self.incidents_admin('Incidencia')),('Justificantes',self.justifications_admin),('Cartas compromiso',self.commitments),('Enviar avisos',self.send_director_message),('Mensajes de docentes',self.inbox_director),('Usuarios',self.users)]
        elif role=='docente': items=[('Pasar lista',self.attendance),('Conductas',lambda:self.incident_form('Conducta')),('Incidencias',lambda:self.incident_form('Incidencia')),('Justificantes',self.justifications_teacher),('Notificaciones recibidas',self.inbox_teacher),('Enviar mensaje al director',self.send_teacher_message)]
        else: items=[('Consulta de alumnos',self.tutor_view)]
        tk.Label(side,text='MENÚ PRINCIPAL',bg=THEME['sidebar'],fg='#B8E6F6',font=('Segoe UI',10,'bold')).pack(anchor='w',padx=17,pady=(22,13))
        for label,fn in items:
            b=tk.Button(side,text=label,command=fn,bg=THEME['sidebar'],fg='white',activebackground=THEME['sidebar_hover'],activeforeground='white',relief='flat',bd=0,anchor='w',padx=17,pady=11,font=('Segoe UI',9,'bold'),cursor='hand2'); b.pack(fill='x',padx=7,pady=2)
        if role=='director': self.summary()
        elif role=='docente': self.attendance()
        else: self.tutor_view()
    def heading(self,title,subtitle=''):
        for w in self.content.winfo_children(): w.destroy()
        tk.Label(self.content,text=title,bg=THEME['background'],fg=THEME['title'],font=('Segoe UI',22,'bold')).pack(anchor='w')
        if subtitle: tk.Label(self.content,text=subtitle,bg=THEME['background'],fg=THEME['subtitle'],font=('Segoe UI',10)).pack(anchor='w',pady=(3,12))
    def table(self,parent,columns,widths=None,height=15):
        frame=tk.Frame(parent,bg=THEME['background']); frame.pack(fill='both',expand=True,pady=5)
        tree=ttk.Treeview(frame,columns=columns,show='headings',height=height)
        for i,c in enumerate(columns): tree.heading(c,text=c); tree.column(c,width=widths[i] if widths else 130,anchor='w')
        sy=ttk.Scrollbar(frame,orient='vertical',command=tree.yview); sx=ttk.Scrollbar(frame,orient='horizontal',command=tree.xview)
        tree.configure(yscrollcommand=sy.set,xscrollcommand=sx.set); tree.grid(row=0,column=0,sticky='nsew'); sy.grid(row=0,column=1,sticky='ns'); sx.grid(row=1,column=0,sticky='ew')
        frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1); return tree
    def card(self,parent,label,value,bg,fg):
        c=tk.Frame(parent,bg=bg,highlightbackground=THEME['border'],highlightthickness=1); c.pack(side='left',fill='x',expand=True,padx=5,pady=5)
        tk.Label(c,text=label.upper(),bg=bg,fg=fg,font=('Segoe UI',9,'bold')).pack(anchor='w',padx=13,pady=(12,2))
        tk.Label(c,text=str(value),bg=bg,fg=fg,font=('Segoe UI',25,'bold')).pack(anchor='w',padx=13,pady=(0,12))
    def scroll_page(self):
        outer=self.content
        canvas=tk.Canvas(outer,bg=THEME['background'],highlightthickness=0)
        bar=ttk.Scrollbar(outer,orient='vertical',command=canvas.yview)
        canvas.configure(yscrollcommand=bar.set)
        bar.pack(side='right',fill='y')
        canvas.pack(side='left',fill='both',expand=True)
        page=tk.Frame(canvas,bg=THEME['background'])
        window=canvas.create_window((0,0),window=page,anchor='nw')
        page.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',lambda e:canvas.itemconfigure(window,width=e.width))
        return page

    def summary(self):
        self.heading('Resumen y asistencia','Indicadores diarios, resumen por grupo y seguimiento de alumnos.')
        self.content=self.scroll_page()
        tk.Label(self.content,text='Fecha de consulta: '+date.today().strftime('%d/%m/%Y'),bg=THEME['background'],fg=THEME['title'],font=('Segoe UI',10,'bold')).pack(anchor='w',pady=5)
        try:
            vals=[]
            for state in ('Presente','Falta','Retardo'): vals.append(query('SELECT COUNT(*) n FROM asistencias WHERE fecha=%s AND estado=%s',(today(),state))[0]['n'])
            row=tk.Frame(self.content,bg=THEME['background']); row.pack(fill='x')
            self.card(row,'Alumnos presentes',vals[0],THEME['card_present_bg'],THEME['card_present_fg'])
            self.card(row,'Faltas',vals[1],THEME['card_absence_bg'],THEME['card_absence_fg'])
            self.card(row,'Retardos',vals[2],THEME['card_late_bg'],THEME['card_late_fg'])
            self.card(row,'Alumnos activos',query('SELECT COUNT(*) n FROM alumnos WHERE activo=1')[0]['n'],THEME['card_neutral_bg'],THEME['card_neutral_fg'])
            tk.Label(self.content,text='Resumen por grupo',bg=THEME['background'],fg=THEME['title'],font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(15,5))
            t=self.table(self.content,['Grupo','Alumnos','Presentes','Faltas','Justificadas','Pendientes','Retardos'],[130,90,95,80,105,100,90],9)
            rows=query("""SELECT g.id,g.nombre,COUNT(DISTINCT a.id) alumnos,
              SUM(s.estado='Presente') presentes,SUM(s.estado='Falta') faltas,
              SUM(s.estado='Falta' AND j.estado='Aprobado') justificadas,
              SUM(s.estado='Falta' AND (j.id IS NULL OR j.estado='Pendiente')) pendientes,
              SUM(s.estado='Retardo') retardos
              FROM grupos g LEFT JOIN alumnos a ON a.grupo_id=g.id AND a.activo=1
              LEFT JOIN asistencias s ON s.alumno_id=a.id AND s.fecha=%s
              LEFT JOIN justificantes j ON j.asistencia_id=s.id WHERE g.activo=1 GROUP BY g.id,g.nombre ORDER BY g.nombre""",(today(),))
            for r in rows: t.insert('','end',values=(r['nombre'],r['alumnos'],r['presentes'] or 0,r['faltas'] or 0,r['justificadas'] or 0,r['pendientes'] or 0,r['retardos'] or 0))
            tk.Label(self.content,text='Seguimiento de alumnos',bg=THEME['background'],fg=THEME['title'],font=('Segoe UI',13,'bold')).pack(anchor='w',pady=(12,4))
            detail=self.table(self.content,['Alumno','Grupo','Matrícula','Tutor','Teléfono','Estado'],[220,100,100,160,120,90],8)
            for r in query("""SELECT a.nombre,g.nombre grupo,a.matricula,a.tutor_nombre,a.tutor_telefono,s.estado
              FROM asistencias s JOIN alumnos a ON a.id=s.alumno_id JOIN grupos g ON g.id=a.grupo_id
              WHERE s.fecha=%s AND s.estado IN ('Falta','Retardo') ORDER BY g.nombre,a.nombre""",(today(),)):
                detail.insert('','end',values=(r['nombre'],r['grupo'],r['matricula'] or '',r['tutor_nombre'] or '',r['tutor_telefono'] or '',r['estado']))
            follow=self.table(self.content,['Alumno','Grupo','Faltas no justificadas','Justificantes pendientes'],[250,130,160,180],7)
            for r in query("""SELECT a.nombre,g.nombre grupo,
              SUM(s.estado='Falta' AND (j.id IS NULL OR j.estado<>'Aprobado')) faltas,
              SUM(j.estado='Pendiente') pendientes FROM alumnos a JOIN grupos g ON g.id=a.grupo_id
              LEFT JOIN asistencias s ON s.alumno_id=a.id LEFT JOIN justificantes j ON j.asistencia_id=s.id
              WHERE a.activo=1 GROUP BY a.id,a.nombre,g.nombre HAVING faltas>0 OR pendientes>0 ORDER BY faltas DESC LIMIT 100"""):
                follow.insert('','end',values=(r['nombre'],r['grupo'],r['faltas'] or 0,r['pendientes'] or 0))
        except Exception as e: tk.Label(self.content,text=str(e),bg=THEME['background'],fg='red').pack(anchor='w')
    def get_group(self):
        rows=query('SELECT * FROM grupos WHERE docente_id=%s AND activo=1 ORDER BY id LIMIT 1',(self.user['id'],)); return rows[0] if rows else None
    def attendance(self):
        self.heading('Pasar lista','Todos los alumnos comienzan como presentes. Selecciona una fila para cambiar su estado.')
        group=self.get_group()
        if not group: tk.Label(self.content,text='Dirección debe asignarte un grupo activo.',bg=THEME['background'],fg='red').pack(anchor='w'); return
        bar=tk.Frame(self.content,bg=THEME['background']); bar.pack(fill='x')
        tk.Label(bar,text=f"{group['nombre']}  •  {date.today().strftime('%d/%m/%Y')}",bg=THEME['background'],fg=THEME['title'],font=('Segoe UI',12,'bold')).pack(side='left')
        students=query('SELECT id,numero_lista,nombre FROM alumnos WHERE grupo_id=%s AND activo=1 ORDER BY numero_lista,nombre',(group['id'],))
        tree=self.table(self.content,['ID','N.º lista','Alumno','Estado'],[60,90,400,140],20); states={}
        tree.tag_configure('Presente',background=THEME['present_row']); tree.tag_configure('Falta',background=THEME['absence_row']); tree.tag_configure('Retardo',background=THEME['late_row'])
        for a in students:
            old=query('SELECT estado FROM asistencias WHERE alumno_id=%s AND fecha=%s',(a['id'],today()))
            state=old[0]['estado'] if old else 'Presente'; states[a['id']]=state
            tree.insert('','end',iid=str(a['id']),values=(a['id'],a['numero_lista'],a['nombre'],state),tags=(state,))
        def setstate(st):
            for iid in tree.selection(): states[int(iid)]=st; tree.set(iid,'Estado',st); tree.item(iid,tags=(st,))
        for st in ('Presente','Falta','Retardo'): self.btn(bar,st,lambda x=st:setstate(x)).pack(side='left',padx=4)
        def all_present():
            for a in students: states[a['id']]='Presente'; tree.set(str(a['id']),'Estado','Presente'); tree.item(str(a['id']),tags=('Presente',))
        self.btn(bar,'Todos presentes',all_present).pack(side='left',padx=4)
        def save():
            try:
                for a in students: execute("""INSERT INTO asistencias(alumno_id,grupo_id,docente_id,fecha,estado) VALUES(%s,%s,%s,%s,%s)
                ON DUPLICATE KEY UPDATE grupo_id=VALUES(grupo_id),docente_id=VALUES(docente_id),estado=VALUES(estado)""",(a['id'],group['id'],self.user['id'],today(),states[a['id']]))
                self.check_absence_alerts(); messagebox.showinfo('Asistencia','La lista se guardó correctamente.'); self.attendance()
            except Exception as e: messagebox.showerror('Asistencia',str(e))
        self.btn(bar,'Guardar asistencia',save,True).pack(side='right')
    def check_absence_alerts(self):
        directors=query("SELECT id FROM usuarios WHERE rol='director' AND activo=1")
        students=query("""SELECT a.id,a.nombre,g.nombre grupo,COUNT(*) total FROM alumnos a JOIN grupos g ON g.id=a.grupo_id
        JOIN asistencias s ON s.alumno_id=a.id AND s.estado='Falta'
        LEFT JOIN justificantes j ON j.asistencia_id=s.id AND j.estado='Aprobado'
        WHERE a.activo=1 AND j.id IS NULL GROUP BY a.id,a.nombre,g.nombre HAVING COUNT(*)>=3""")
        for a in students:
            for d in directors:
                exists=query("SELECT id FROM notificaciones WHERE destinatario_id=%s AND tipo='TRES_FALTAS' AND referencia_id=%s AND leida=0",(d['id'],a['id']))
                if not exists: execute("INSERT INTO notificaciones(destinatario_id,tipo,titulo,mensaje,referencia_id,prioridad) VALUES(%s,'TRES_FALTAS',%s,%s,%s,'alta')",(d['id'],'Alerta: tres faltas no justificadas',f"{a['nombre']} ({a['grupo']}) acumula {a['total']} faltas sin justificar.",a['id']))
    def students(self):
        self.heading('Alumnos','Alta, edición individual e importación masiva desde Excel.')
        self.content=self.scroll_page()
        bar=tk.Frame(self.content,bg=THEME['background']); bar.pack(fill='x',pady=5)
        self.btn(bar,'Agregar alumno',lambda:self.student_form(),True).pack(side='left',padx=3)
        self.btn(bar,'Editar seleccionado',self.edit_selected_student).pack(side='left',padx=3)
        self.btn(bar,'Desactivar seleccionado',self.deactivate_student).pack(side='left',padx=3)
        groups=query('SELECT id,nombre FROM grupos WHERE activo=1 ORDER BY nombre')
        for g in groups:
            tk.Label(self.content,text='Grupo '+g['nombre'],bg=THEME['background'],fg=THEME['title'],font=('Segoe UI',13,'bold')).pack(anchor='w',pady=8)
            self.btn(self.content,'Importar Excel - '+g['nombre'],lambda group=g:self.import_excel(group)).pack(anchor='w')
            t=self.table(self.content,['ID','N.º','Alumno','Grupo','Matrícula','Tutor','Teléfono','Activo'],[55,55,210,100,100,180,110,70],6)
            t.master.pack_configure(expand=False)
            t.bind('<<TreeviewSelect>>',lambda e,tree=t:setattr(self,'student_table',tree))
            for r in query('SELECT a.*,g.nombre grupo FROM alumnos a JOIN grupos g ON g.id=a.grupo_id WHERE a.grupo_id=%s ORDER BY a.numero_lista,a.nombre',(g['id'],)):
                t.insert('','end',iid=str(r['id']),values=(r['id'],r['numero_lista'],r['nombre'],r['grupo'],r.get('matricula') or '',r.get('tutor_nombre') or '',r.get('tutor_telefono') or '','Sí' if r['activo'] else 'No'))
            self.student_table=t
    def selected_id(self,attr):
        t=getattr(self,attr,None)
        if not t or not t.selection(): messagebox.showwarning('Selecciona','Selecciona un registro de la tabla.'); return None
        return int(t.selection()[0])
    def deactivate_student(self):
        sid=self.selected_id('student_table')
        if sid and messagebox.askyesno('Confirmar','¿Deseas desactivar este alumno?'): execute('UPDATE alumnos SET activo=0 WHERE id=%s',(sid,)); self.students()
    def edit_selected_student(self):
        sid=self.selected_id('student_table')
        if sid: self.student_form(sid)
    def student_form(self,sid=None):
        groups=query('SELECT id,nombre FROM grupos WHERE activo=1 ORDER BY nombre')
        if not groups: messagebox.showwarning('Grupos','Primero registra un grupo.'); return
        old=(query('SELECT * FROM alumnos WHERE id=%s',(sid,)) or [None])[0] if sid else None
        win=tk.Toplevel(self); win.title('Editar alumno' if sid else 'Agregar alumno'); win.geometry('520x650'); win.configure(bg='white'); win.grab_set()
        fields={}
        specs=[('Nombre completo','nombre'),('Número de lista','numero_lista'),('Matrícula','matricula'),('Fecha nacimiento (AAAA-MM-DD)','fecha_nacimiento'),('Nombre del tutor','tutor_nombre'),('Teléfono del tutor','tutor_telefono'),('Correo del tutor','tutor_correo'),('Domicilio','domicilio')]
        for label,key in specs:
            tk.Label(win,text=label,bg='white',fg=THEME['title'],font=('Segoe UI',9,'bold')).pack(anchor='w',padx=18,pady=(9,1))
            e=tk.Entry(win,font=('Segoe UI',10)); e.pack(fill='x',padx=18,ipady=5); e.insert(0,str((old or {}).get(key) or '')); fields[key]=e
        tk.Label(win,text='Grupo',bg='white',fg=THEME['title'],font=('Segoe UI',9,'bold')).pack(anchor='w',padx=18,pady=(9,1))
        gvar=tk.StringVar(value=next((g['nombre'] for g in groups if old and g['id']==old['grupo_id']),groups[0]['nombre']))
        ttk.Combobox(win,textvariable=gvar,values=[g['nombre'] for g in groups],state='readonly').pack(fill='x',padx=18)
        tk.Label(win,text='Observaciones',bg='white',fg=THEME['title'],font=('Segoe UI',9,'bold')).pack(anchor='w',padx=18,pady=(9,1))
        obs=tk.Text(win,height=3); obs.pack(fill='x',padx=18); obs.insert('1.0',(old or {}).get('observaciones') or '')
        def save():
            try:
                name=fields['nombre'].get().strip(); num=int(fields['numero_lista'].get() or 0); g=next(x for x in groups if x['nombre']==gvar.get())
                if not name: raise ValueError('El nombre es obligatorio.')
                if query('SELECT id FROM alumnos WHERE grupo_id=%s AND numero_lista=%s AND id<>%s',(g['id'],num,sid or 0)): raise ValueError('Ese número de lista ya existe en el grupo.')
                vals=(g['id'],num,name,fields['matricula'].get().strip() or None,fields['fecha_nacimiento'].get().strip() or None,fields['tutor_nombre'].get().strip() or None,fields['tutor_telefono'].get().strip() or None,fields['tutor_correo'].get().strip() or None,fields['domicilio'].get().strip() or None,obs.get('1.0','end').strip() or None)
                if sid: execute('UPDATE alumnos SET grupo_id=%s,numero_lista=%s,nombre=%s,matricula=%s,fecha_nacimiento=%s,tutor_nombre=%s,tutor_telefono=%s,tutor_correo=%s,domicilio=%s,observaciones=%s WHERE id=%s',vals+(sid,))
                else: execute('INSERT INTO alumnos(grupo_id,numero_lista,nombre,matricula,fecha_nacimiento,tutor_nombre,tutor_telefono,tutor_correo,domicilio,observaciones) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)',vals)
                win.destroy(); self.students()
            except Exception as e: messagebox.showerror('Alumno',str(e))
        self.btn(win,'Guardar alumno',save,True).pack(pady=15)
    def import_excel(self, group=None):
        path=filedialog.askopenfilename(filetypes=[('Libro de Excel','*.xlsx')])
        if not path:return
        try:
            from openpyxl import load_workbook
            ws=load_workbook(path,data_only=True,read_only=True).active
            headers=[str(c or '').strip().lower() for c in next(ws.iter_rows(min_row=1,max_row=1,values_only=True))]
            required={'nombre','grupo'}
            if not required.issubset(set(headers)): raise ValueError('El Excel debe tener encabezados: nombre y grupo. Opcionales: numero_lista, matricula, tutor_nombre, tutor_telefono, tutor_correo, domicilio.')
            idx={h:i for i,h in enumerate(headers)}; groups={r['nombre'].strip().lower():r['id'] for r in query('SELECT id,nombre FROM grupos WHERE activo=1')}
            prepared=[]; errors=[]; seen=set(); nextnum={g['id']:query('SELECT COALESCE(MAX(numero_lista),0)+1 n FROM alumnos WHERE grupo_id=%s',(g['id'],))[0]['n'] for g in query('SELECT id FROM grupos WHERE activo=1')}
            for line,row in enumerate(ws.iter_rows(min_row=2,values_only=True),start=2):
                if not row or not any(v is not None and str(v).strip() for v in row): continue
                get=lambda k: row[idx[k]] if k in idx and idx[k]<len(row) else None
                name=str(get('nombre') or '').strip(); gn=str(get('grupo') or '').strip().lower()
                if not name: errors.append(f'Fila {line}: falta nombre.'); continue
                if gn not in groups: errors.append(f'Fila {line}: grupo no registrado ({get("grupo")}).'); continue
                gid=groups[gn]; number=get('numero_lista')
                if group and gid!=group['id']: errors.append(f'Fila {line}: el grupo no corresponde a esta tabla.'); continue
                try: number=int(number) if number not in (None,'') else int(nextnum[gid])
                except Exception: errors.append(f'Fila {line}: numero_lista debe ser numérico.'); continue
                key=(gid,number)
                if key in seen or query('SELECT id FROM alumnos WHERE grupo_id=%s AND numero_lista=%s',(gid,number)): errors.append(f'Fila {line}: número de lista {number} duplicado en el grupo.'); continue
                seen.add(key); nextnum[gid]=max(nextnum[gid],number+1)
                prepared.append((gid,number,name,str(get('matricula') or '').strip() or None,str(get('tutor_nombre') or '').strip() or None,str(get('tutor_telefono') or '').strip() or None,str(get('tutor_correo') or '').strip() or None,str(get('domicilio') or '').strip() or None))
            if not prepared: raise ValueError('No hay filas válidas para importar.\n'+'\n'.join(errors[:15]))
            preview=tk.Toplevel(self); preview.title('Vista previa de importación'); preview.geometry('850x560'); preview.grab_set()
            tk.Label(preview,text=f'Registros válidos: {len(prepared)}   |   Filas con errores: {len(errors)}',font=('Segoe UI',12,'bold')).pack(anchor='w',padx=12,pady=10)
            t=self.table(preview,['Grupo ID','N.º','Nombre','Matrícula','Tutor','Teléfono'],[80,55,230,100,170,110],12)
            for r in prepared[:300]: t.insert('','end',values=r[:6])
            if errors:
                box=tk.Text(preview,height=5); box.pack(fill='x',padx=12); box.insert('1.0','\n'.join(errors[:30])); box.configure(state='disabled')
            def confirm():
                try:
                    count=executemany('INSERT INTO alumnos(grupo_id,numero_lista,nombre,matricula,tutor_nombre,tutor_telefono,tutor_correo,domicilio) VALUES(%s,%s,%s,%s,%s,%s,%s,%s)',prepared)
                    preview.destroy(); messagebox.showinfo('Importación',f'Importación terminada. Registros procesados: {count}. Filas omitidas: {len(errors)}.'); self.students()
                except Exception as e: messagebox.showerror('Importación',str(e))
            self.btn(preview,'Confirmar importación',confirm,True).pack(anchor='e',padx=12,pady=10)
        except ImportError: messagebox.showerror('Excel','Falta openpyxl. Instálalo con: py -m pip install openpyxl')
        except Exception as e: messagebox.showerror('Excel',str(e))
    def groups(self):
        self.heading('Docentes y grupos','Administra los grupos y edita la información general de cada docente.')
        bar=tk.Frame(self.content,bg=THEME['background']); bar.pack(fill='x')
        self.btn(bar,'Nuevo grupo',self.add_group,True).pack(side='left',padx=3); self.btn(bar,'Editar grupo',self.edit_group).pack(side='left',padx=3); self.btn(bar,'Editar docente seleccionado',self.edit_teacher).pack(side='left',padx=3); self.btn(bar,'Nuevo docente',self.add_teacher).pack(side='left',padx=3)
        tk.Label(self.content,text='Grupos y asignación',bg=THEME['background'],fg=THEME['title'],font=('Segoe UI',12,'bold')).pack(anchor='w',pady=(12,3))
        gt=self.table(self.content,['ID','Grupo','Docente asignado','Alumnos','Activo'],[55,150,230,100,80],7); self.group_table=gt
        for r in query('SELECT g.id,g.nombre,u.nombre docente,(SELECT COUNT(*) FROM alumnos a WHERE a.grupo_id=g.id AND a.activo=1) alumnos,g.activo FROM grupos g LEFT JOIN usuarios u ON u.id=g.docente_id ORDER BY g.nombre'):
            gt.insert('','end',iid=str(r['id']),values=(r['id'],r['nombre'],r['docente'] or 'Sin asignar',r['alumnos'],'Sí' if r['activo'] else 'No'))
        tk.Label(self.content,text='Docentes',bg=THEME['background'],fg=THEME['title'],font=('Segoe UI',12,'bold')).pack(anchor='w',pady=(12,3))
        dt=self.table(self.content,['ID','Nombre','Usuario','Grupo','Teléfono','Correo','Activo'],[55,180,110,120,110,180,70],7); self.teacher_table=dt
        for r in query("SELECT u.id,u.nombre,u.usuario,u.telefono,u.correo,u.activo,g.nombre grupo FROM usuarios u LEFT JOIN grupos g ON g.docente_id=u.id WHERE u.rol='docente' ORDER BY u.nombre"):
            dt.insert('','end',iid=str(r['id']),values=(r['id'],r['nombre'],r['usuario'],r['grupo'] or 'Sin grupo',r['telefono'] or '',r['correo'] or '','Sí' if r['activo'] else 'No'))
    def add_group(self):
        name=simpledialog.askstring('Nuevo grupo','Nombre del grupo (ej. 1° A):',parent=self)
        if name and name.strip():
            try: execute('INSERT INTO grupos(nombre) VALUES(%s)',(name.strip(),)); self.groups()
            except Exception as e: messagebox.showerror('Grupo',str(e))
    def edit_group(self):
        gid=self.selected_id('group_table')
        if not gid:return
        g=query('SELECT * FROM grupos WHERE id=%s',(gid,))[0]; name=simpledialog.askstring('Editar grupo','Nombre del grupo:',initialvalue=g['nombre'],parent=self)
        if name and name.strip():
            try: execute('UPDATE grupos SET nombre=%s WHERE id=%s',(name.strip(),gid)); self.groups()
            except Exception as e: messagebox.showerror('Grupo',str(e))
    def teacher_form(self,uid=None):
        old=(query('SELECT * FROM usuarios WHERE id=%s',(uid,)) or [None])[0] if uid else None
        groups=query('SELECT id,nombre FROM grupos WHERE activo=1 ORDER BY nombre'); win=tk.Toplevel(self); win.title('Datos del docente'); win.geometry('470x470'); win.grab_set(); fields={}
        for label,key in [('Nombre completo','nombre'),('Nombre de usuario','usuario'),('Teléfono','telefono'),('Correo','correo'),('Dirección','direccion')]:
            tk.Label(win,text=label).pack(anchor='w',padx=18,pady=(9,1)); e=tk.Entry(win); e.pack(fill='x',padx=18); e.insert(0,(old or {}).get(key) or ''); fields[key]=e
        tk.Label(win,text='Contraseña inicial (dejar vacía para conservar)').pack(anchor='w',padx=18,pady=(9,1)); pw=tk.Entry(win,show='*'); pw.pack(fill='x',padx=18)
        tk.Label(win,text='Grupo asignado').pack(anchor='w',padx=18,pady=(9,1)); current=next((g['nombre'] for g in groups if old and query('SELECT id FROM grupos WHERE id=%s AND docente_id=%s',(g['id'],uid))),groups[0]['nombre'] if groups else ''); gv=tk.StringVar(value=current); ttk.Combobox(win,textvariable=gv,values=[g['nombre'] for g in groups],state='readonly').pack(fill='x',padx=18)
        def save():
            try:
                name=fields['nombre'].get().strip(); user=fields['usuario'].get().strip()
                if not name or not user: raise ValueError('Nombre y usuario son obligatorios.')
                if uid:
                    sql='UPDATE usuarios SET nombre=%s,usuario=%s,telefono=%s,correo=%s,direccion=%s WHERE id=%s'; execute(sql,(name,user,fields['telefono'].get().strip(),fields['correo'].get().strip(),fields['direccion'].get().strip(),uid))
                    if pw.get().strip(): execute('UPDATE usuarios SET password_hash=%s WHERE id=%s',(hash_pw(pw.get().strip()),uid))
                    execute('UPDATE grupos SET docente_id=NULL WHERE docente_id=%s',(uid,))
                    if gv.get(): g=next(x for x in groups if x['nombre']==gv.get()); execute('UPDATE grupos SET docente_id=%s WHERE id=%s',(uid,g['id']))
                else:
                    if not pw.get().strip(): raise ValueError('Define una contraseña inicial.')
                    new=execute("INSERT INTO usuarios(nombre,usuario,password_hash,rol,telefono,correo,direccion) VALUES(%s,%s,%s,'docente',%s,%s,%s)",(name,user,hash_pw(pw.get().strip()),fields['telefono'].get().strip(),fields['correo'].get().strip(),fields['direccion'].get().strip()))
                    if gv.get(): g=next(x for x in groups if x['nombre']==gv.get()); execute('UPDATE grupos SET docente_id=%s WHERE id=%s',(new,g['id']))
                win.destroy(); self.groups()
            except Exception as e: messagebox.showerror('Docente',str(e))
        self.btn(win,'Guardar docente',save,True).pack(pady=15)
    def edit_teacher(self):
        uid=self.selected_id('teacher_table')
        if uid:self.teacher_form(uid)
    def add_teacher(self): self.teacher_form()
    def incident_form(self,kind):
        group=self.get_group()
        self.heading('Registrar '+kind.lower(),f'El reporte se enviará a dirección y quedará guardado en el historial de {kind.lower()}.')
        if not group: tk.Label(self.content,text='No tienes un grupo asignado.',bg=THEME['background']).pack(); return
        students=query('SELECT id,nombre,grupo_id FROM alumnos WHERE grupo_id=%s AND activo=1 ORDER BY numero_lista,nombre',(group['id'],))
        if not students: tk.Label(self.content,text='No hay alumnos activos en tu grupo.',bg=THEME['background']).pack(); return
        form=tk.Frame(self.content,bg='white'); form.pack(fill='x',pady=10); opts={f"{a['nombre']} · {group['nombre']}":a for a in students}; sv=tk.StringVar(value=next(iter(opts)))
        tk.Label(form,text='Alumno',bg='white').pack(anchor='w',padx=14,pady=(12,2)); ttk.Combobox(form,textvariable=sv,values=list(opts),state='readonly').pack(fill='x',padx=14)
        tk.Label(form,text='Título',bg='white').pack(anchor='w',padx=14,pady=(9,2)); title=tk.Entry(form); title.pack(fill='x',padx=14)
        tk.Label(form,text='Gravedad',bg='white').pack(anchor='w',padx=14,pady=(9,2)); sev=tk.StringVar(value='Leve'); ttk.Combobox(form,textvariable=sev,values=['Leve','Media','Grave'],state='readonly').pack(fill='x',padx=14)
        tk.Label(form,text='Descripción',bg='white').pack(anchor='w',padx=14,pady=(9,2)); desc=tk.Text(form,height=5); desc.pack(fill='x',padx=14)
        tk.Label(form,text='Medida o seguimiento (opcional)',bg='white').pack(anchor='w',padx=14,pady=(9,2)); measure=tk.Entry(form); measure.pack(fill='x',padx=14)
        def save():
            a=opts[sv.get()]; ti=title.get().strip(); de=desc.get('1.0','end').strip()
            if not ti or not de: messagebox.showwarning('Datos','Completa el título y la descripción.'); return
            try:
                iid=execute('INSERT INTO incidencias(alumno_id,grupo_id,docente_id,tipo,titulo,gravedad,descripcion,medida,fecha) VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s)',(a['id'],group['id'],self.user['id'],kind,ti,sev.get(),de,measure.get().strip(),today()))
                directors=query("SELECT id FROM usuarios WHERE rol='director' AND activo=1")
                priority='urgente' if sev.get()=='Grave' else 'alta' if sev.get()=='Media' else 'normal'
                for d in directors: execute('INSERT INTO notificaciones(destinatario_id,emisor_id,tipo,titulo,mensaje,referencia_id,prioridad) VALUES(%s,%s,%s,%s,%s,%s,%s)',(d['id'],self.user['id'],kind.upper(),f'{kind}: {ti}',f"Alumno: {a['nombre']}\nGrupo: {group['nombre']}\nDocente: {self.user['nombre']}\nGravedad: {sev.get()}\n\n{de}\n\nMedida: {measure.get().strip() or 'Sin especificar'}",iid,priority))
                messagebox.showinfo(kind,'Reporte guardado y enviado a dirección.'); self.incident_form(kind)
            except Exception as e: messagebox.showerror(kind,str(e))
        self.btn(form,'Enviar reporte a dirección',save,True).pack(anchor='e',padx=14,pady=12)
    def incidents_admin(self,kind):
        self.heading(kind+' recibidas','Reportes enviados por los docentes. Estos registros también permanecen en su historial correspondiente.')
        t=self.table(self.content,['Fecha','Alumno','Grupo','Docente','Gravedad','Título','Descripción','Medida'],[95,160,90,150,85,150,300,200],16)
        for r in query('SELECT i.id,i.fecha,a.nombre alumno,g.nombre grupo,u.nombre docente,i.gravedad,i.titulo,i.descripcion,i.medida FROM incidencias i JOIN alumnos a ON a.id=i.alumno_id JOIN grupos g ON g.id=i.grupo_id JOIN usuarios u ON u.id=i.docente_id WHERE i.tipo=%s ORDER BY i.id DESC LIMIT 400',(kind,)):
            t.insert('','end',iid=str(r['id']),values=(fmtdate(r['fecha']),r['alumno'],r['grupo'],r['docente'],r['gravedad'],r['titulo'],r['descripcion'],r['medida'] or ''))
    def justifications_teacher(self):
        self.heading('Justificantes','Registra un justificante para una falta. Dirección revisará y aprobará o rechazará la solicitud.')
        group=self.get_group()
        if not group: return
        rows=query("SELECT s.id,s.fecha,a.nombre FROM asistencias s JOIN alumnos a ON a.id=s.alumno_id WHERE s.grupo_id=%s AND s.estado='Falta' ORDER BY s.fecha DESC,a.nombre LIMIT 250",(group['id'],))
        opts={f"{r['nombre']} · {fmtdate(r['fecha'])}":r['id'] for r in rows}
        if opts:
            form=tk.Frame(self.content,bg='white'); form.pack(fill='x',pady=8); sv=tk.StringVar(value=next(iter(opts))); ttk.Combobox(form,textvariable=sv,values=list(opts),state='readonly').pack(fill='x',padx=12,pady=8)
            reason=tk.Entry(form); reason.pack(fill='x',padx=12,pady=8); reason.insert(0,'Motivo del justificante')
            def save():
                aid=opts[sv.get()]; why=reason.get().strip()
                if not why or why=='Motivo del justificante': messagebox.showwarning('Datos','Escribe el motivo.'); return
                try:
                    old=query('SELECT id FROM justificantes WHERE asistencia_id=%s',(aid,))
                    if old: execute("UPDATE justificantes SET motivo=%s,estado='Pendiente',respuesta=NULL,revisado_en=NULL WHERE asistencia_id=%s",(why,aid))
                    else: execute('INSERT INTO justificantes(asistencia_id,motivo) VALUES(%s,%s)',(aid,why))
                    messagebox.showinfo('Justificante','Justificante enviado a dirección.'); self.justifications_teacher()
                except Exception as e: messagebox.showerror('Justificante',str(e))
            self.btn(form,'Enviar justificante',save,True).pack(anchor='e',padx=12,pady=8)
        else: tk.Label(self.content,text='No hay faltas registradas para justificar.',bg=THEME['background']).pack(anchor='w')
        t=self.table(self.content,['Fecha','Alumno','Motivo','Estado','Respuesta'],[110,180,320,120,250],10)
        for r in query('SELECT s.fecha,a.nombre,j.motivo,j.estado,j.respuesta FROM justificantes j JOIN asistencias s ON s.id=j.asistencia_id JOIN alumnos a ON a.id=s.alumno_id WHERE s.grupo_id=%s ORDER BY j.id DESC',(group['id'],)):
            t.insert('','end',values=(fmtdate(r['fecha']),r['nombre'],r['motivo'],r['estado'],r['respuesta'] or ''))
    def justifications_admin(self):
        self.heading('Justificantes recibidos','Revisa las solicitudes de justificantes enviadas por los docentes.')
        t=self.table(self.content,['ID','Fecha','Alumno','Grupo','Docente','Motivo','Estado','Respuesta'],[50,95,160,90,150,260,110,180],15)
        for r in query('SELECT j.id,s.fecha,a.nombre alumno,g.nombre grupo,u.nombre docente,j.motivo,j.estado,j.respuesta FROM justificantes j JOIN asistencias s ON s.id=j.asistencia_id JOIN alumnos a ON a.id=s.alumno_id JOIN grupos g ON g.id=s.grupo_id JOIN usuarios u ON u.id=s.docente_id ORDER BY j.id DESC'):
            t.insert('','end',iid=str(r['id']),values=(r['id'],fmtdate(r['fecha']),r['alumno'],r['grupo'],r['docente'],r['motivo'],r['estado'],r['respuesta'] or ''))
        bar=tk.Frame(self.content,bg=THEME['background']); bar.pack(fill='x')
        def review(state):
            jid=self.selected_id_from(t)
            if not jid:return
            response=simpledialog.askstring('Respuesta','Respuesta para el docente (opcional):',parent=self) or ''
            execute('UPDATE justificantes SET estado=%s,respuesta=%s,revisado_en=NOW() WHERE id=%s',(state,response,jid))
            row=query('SELECT u.id destinatario,a.nombre alumno FROM justificantes j JOIN asistencias s ON s.id=j.asistencia_id JOIN alumnos a ON a.id=s.alumno_id JOIN usuarios u ON u.id=s.docente_id WHERE j.id=%s',(jid,))[0]
            execute('INSERT INTO notificaciones(destinatario_id,emisor_id,tipo,titulo,mensaje,prioridad) VALUES(%s,%s,\'JUSTIFICANTE\',%s,%s,\'normal\')',(row['destinatario'],self.user['id'],f'Justificante {state.lower()}',f'El justificante de {row["alumno"]} fue {state.lower()}. {response}'))
            self.justifications_admin()
        self.btn(bar,'Aprobar',lambda:review('Aprobado'),True).pack(side='left',padx=4); self.btn(bar,'Rechazar',lambda:review('Rechazado')).pack(side='left',padx=4)
    def selected_id_from(self,t):
        if not t.selection(): messagebox.showwarning('Selecciona','Selecciona un registro.'); return None
        return int(t.selection()[0])
    def commitments(self):
        self.heading('Cartas compromiso','Crea, consulta e imprime cartas compromiso para dar seguimiento a los acuerdos escolares.')
        students=query('SELECT a.id,a.nombre,g.id grupo_id,g.nombre grupo FROM alumnos a JOIN grupos g ON g.id=a.grupo_id WHERE a.activo=1 ORDER BY g.nombre,a.nombre')
        if not students: tk.Label(self.content,text='No hay alumnos activos.',bg=THEME['background']).pack(); return
        form=tk.Frame(self.content,bg='white'); form.pack(fill='x',pady=7); opts={f"{a['nombre']} · {a['grupo']}":a for a in students}; sv=tk.StringVar(value=next(iter(opts))); ttk.Combobox(form,textvariable=sv,values=list(opts),state='readonly').pack(fill='x',padx=12,pady=7)
        title=tk.Entry(form); title.pack(fill='x',padx=12,pady=4); title.insert(0,'Título de la carta')
        reason=tk.Entry(form); reason.pack(fill='x',padx=12,pady=4); reason.insert(0,'Motivo')
        commit=tk.Text(form,height=3); commit.pack(fill='x',padx=12,pady=4); commit.insert('1.0','Compromisos establecidos')
        follow=tk.Entry(form); follow.pack(fill='x',padx=12,pady=4); follow.insert(0,'Fecha de seguimiento AAAA-MM-DD (opcional)')
        def save():
            a=opts[sv.get()]; why=reason.get().strip(); commitments=commit.get('1.0','end').strip(); ti=title.get().strip() or 'Carta compromiso'; f=follow.get().strip()
            if not why or why=='Motivo' or not commitments or commitments=='Compromisos establecidos': messagebox.showwarning('Datos','Completa el motivo y los compromisos.'); return
            if f and f.startswith('Fecha de seguimiento'): f=''
            try:
                if f: datetime.strptime(f,'%Y-%m-%d')
                execute('INSERT INTO cartas_compromiso(alumno_id,grupo_id,director_id,fecha,titulo,motivo,compromisos,fecha_seguimiento) VALUES(%s,%s,%s,%s,%s,%s,%s,%s)',(a['id'],a['grupo_id'],self.user['id'],today(),ti,why,commitments,f or None))
                messagebox.showinfo('Carta','Carta compromiso registrada.'); self.commitments()
            except Exception as e: messagebox.showerror('Carta',str(e))
        self.btn(form,'Guardar carta compromiso',save,True).pack(anchor='e',padx=12,pady=8)
        t=self.table(self.content,['ID','Fecha','Alumno','Grupo','Título','Motivo','Compromisos','Seguimiento','Estado'],[45,90,150,90,130,180,220,100,110],10); self.commit_table=t
        for r in query('SELECT c.*,a.nombre alumno,g.nombre grupo FROM cartas_compromiso c JOIN alumnos a ON a.id=c.alumno_id JOIN grupos g ON g.id=c.grupo_id ORDER BY c.id DESC LIMIT 300'):
            t.insert('','end',iid=str(r['id']),values=(r['id'],fmtdate(r['fecha']),r['alumno'],r['grupo'],r['titulo'],r['motivo'],r['compromisos'],fmtdate(r['fecha_seguimiento']) if r['fecha_seguimiento'] else '',r['estado']))
        bar=tk.Frame(self.content,bg=THEME['background']); bar.pack(fill='x')
        self.btn(bar,'Imprimir / guardar PDF',lambda:self.print_commitment(t)).pack(side='left',padx=3)
        self.btn(bar,'Cambiar estado',lambda:self.change_commitment_state(t)).pack(side='left',padx=3)
    def print_commitment(self,t):
        cid=self.selected_id_from(t)
        if not cid:return
        r=query('SELECT c.*,a.nombre alumno,g.nombre grupo,u.nombre director FROM cartas_compromiso c JOIN alumnos a ON a.id=c.alumno_id JOIN grupos g ON g.id=c.grupo_id JOIN usuarios u ON u.id=c.director_id WHERE c.id=%s',(cid,))[0]
        path=filedialog.asksaveasfilename(defaultextension='.txt',filetypes=[('Documento de texto','*.txt'),('PDF','*.pdf')],initialfile=f'Carta_compromiso_{cid}.txt')
        if not path:return
        content=f"PRIMARIA NEZAHUALCÓYOTL\nCARTA COMPROMISO\n\nFecha: {fmtdate(r['fecha'])}\nAlumno: {r['alumno']}\nGrupo: {r['grupo']}\nTítulo: {r['titulo']}\n\nMotivo:\n{r['motivo']}\n\nCompromisos:\n{r['compromisos']}\n\nFecha de seguimiento: {fmtdate(r['fecha_seguimiento']) if r['fecha_seguimiento'] else 'No definida'}\nEstado: {r['estado']}\n\nDirección: {r['director']}\n\nFirma del tutor: __________________________\nFirma de dirección: _______________________\n"
        try:
            if path.lower().endswith('.pdf'):
                try:
                    from reportlab.lib.pagesizes import letter
                    from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer
                    from reportlab.lib.styles import getSampleStyleSheet
                    from reportlab.lib.units import inch
                    doc=SimpleDocTemplate(path,pagesize=letter); styles=getSampleStyleSheet(); story=[]
                    for line in content.splitlines(): story.append(Paragraph(line.replace('&','&amp;') or '&nbsp;',styles['Normal'])); story.append(Spacer(1,0.12*inch))
                    doc.build(story)
                except ImportError: raise RuntimeError('Para generar PDF instala reportlab: py -m pip install reportlab. También puedes guardar como TXT.')
            else:
                with open(path,'w',encoding='utf-8') as f:f.write(content)
            messagebox.showinfo('Carta','Documento guardado correctamente.')
        except Exception as e: messagebox.showerror('Carta',str(e))
    def change_commitment_state(self,t):
        cid=self.selected_id_from(t)
        if not cid:return
        state=simpledialog.askstring('Estado','Escribe: Pendiente, En seguimiento, Cumplida o Cerrada',parent=self)
        if state in ('Pendiente','En seguimiento','Cumplida','Cerrada'): execute('UPDATE cartas_compromiso SET estado=%s WHERE id=%s',(state,cid)); self.commitments()
        elif state: messagebox.showwarning('Estado','Estado no válido.')
    def send_director_message(self): self.message_form('DIRECTOR_DOCENTE')
    def send_teacher_message(self): self.message_form('DOCENTE_DIRECTOR')
    def message_form(self,kind):
        director=kind=='DIRECTOR_DOCENTE'; self.heading('Enviar avisos a docentes' if director else 'Enviar mensaje al director','Los mensajes quedan guardados y disponibles aunque el destinatario no esté conectado.')
        form=tk.Frame(self.content,bg='white'); form.pack(fill='x',pady=10)
        if director:
            recipients=query("SELECT id,nombre FROM usuarios WHERE rol='docente' AND activo=1 ORDER BY nombre"); opts={'Todos los docentes':None}; opts.update({r['nombre']:r['id'] for r in recipients}); label='Destinatario'
        else:
            recipients=query("SELECT id,nombre FROM usuarios WHERE rol='director' AND activo=1 ORDER BY nombre"); opts={r['nombre']:r['id'] for r in recipients}; label='Dirección'
        if not opts: tk.Label(self.content,text='No hay destinatarios activos.',bg=THEME['background']).pack(); return
        sv=tk.StringVar(value=next(iter(opts))); tk.Label(form,text=label,bg='white').pack(anchor='w',padx=14,pady=(12,2)); ttk.Combobox(form,textvariable=sv,values=list(opts),state='readonly').pack(fill='x',padx=14)
        tk.Label(form,text='Asunto',bg='white').pack(anchor='w',padx=14,pady=(10,2)); title=tk.Entry(form); title.pack(fill='x',padx=14)
        tk.Label(form,text='Mensaje',bg='white').pack(anchor='w',padx=14,pady=(10,2)); body=tk.Text(form,height=6); body.pack(fill='x',padx=14)
        def send():
            ti=title.get().strip(); msg=body.get('1.0','end').strip()
            if not ti or not msg: messagebox.showwarning('Datos','Completa asunto y mensaje.'); return
            try:
                targets=recipients if director and opts[sv.get()] is None else [next(x for x in recipients if x['id']==opts[sv.get()])]
                for target in targets: execute('INSERT INTO mensajes(remitente_id,destinatario_id,tipo,titulo,mensaje) VALUES(%s,%s,%s,%s,%s)',(self.user['id'],target['id'],kind,ti,msg))
                messagebox.showinfo('Mensaje',f'Mensaje enviado a {len(targets)} destinatario(s).'); self.message_form(kind)
            except Exception as e: messagebox.showerror('Mensaje',str(e))
        self.btn(form,'Enviar mensaje',send,True).pack(anchor='e',padx=14,pady=12)
    def inbox_teacher(self): self.message_inbox('DIRECTOR_DOCENTE','Notificaciones recibidas del director')
    def inbox_director(self): self.message_inbox('DOCENTE_DIRECTOR','Mensajes recibidos de docentes')
    def message_inbox(self,kind,title):
        self.heading(title,'Bandeja separada de los reportes de conducta, incidencias y justificantes.')
        t=self.table(self.content,['ID','Fecha','Remitente','Asunto','Mensaje','Estado'],[50,130,170,190,400,90],16)
        for r in query('SELECT m.*,u.nombre remitente FROM mensajes m JOIN usuarios u ON u.id=m.remitente_id WHERE m.destinatario_id=%s AND m.tipo=%s ORDER BY m.leido,m.creado_en DESC',(self.user['id'],kind)):
            t.insert('','end',iid=str(r['id']),values=(r['id'],r['creado_en'],r['remitente'],r['titulo'],r['mensaje'],'Leído' if r['leido'] else 'Nuevo'))
        def mark():
            mid=self.selected_id_from(t)
            if mid: execute('UPDATE mensajes SET leido=1 WHERE id=%s AND destinatario_id=%s',(mid,self.user['id'])); self.message_inbox(kind,title)
        self.btn(self.content,'Marcar seleccionado como leído',mark).pack(anchor='e',pady=5)
    def users(self):
        self.heading('Usuarios','Administración de cuentas, roles y contraseñas.')
        bar=tk.Frame(self.content,bg=THEME['background']); bar.pack(fill='x'); self.btn(bar,'Crear usuario',self.add_user,True).pack(side='left',padx=3); self.btn(bar,'Cambiar contraseña',self.change_password).pack(side='left',padx=3); self.btn(bar,'Desactivar cuenta',self.deactivate_user).pack(side='left',padx=3)
        t=self.table(self.content,['ID','Usuario','Nombre','Rol','Teléfono','Correo','Activo'],[55,120,200,100,120,200,70],16); self.user_table=t
        for r in query('SELECT id,usuario,nombre,rol,telefono,correo,activo FROM usuarios ORDER BY rol,nombre'):
            t.insert('','end',iid=str(r['id']),values=(r['id'],r['usuario'],r['nombre'],r['rol'],r['telefono'] or '',r['correo'] or '','Sí' if r['activo'] else 'No'))
    def add_user(self):
        win=tk.Toplevel(self); win.title('Crear usuario'); win.geometry('430x390'); win.grab_set(); fields=[]
        for label,secret in [('Nombre completo',False),('Usuario',False),('Contraseña inicial',True),('Teléfono',False),('Correo',False)]:
            tk.Label(win,text=label).pack(anchor='w',padx=18,pady=(9,1)); e=tk.Entry(win,show='*' if secret else ''); e.pack(fill='x',padx=18); fields.append(e)
        role=tk.StringVar(value='docente'); ttk.Combobox(win,textvariable=role,values=['director','docente','tutor'],state='readonly').pack(fill='x',padx=18,pady=12)
        def save():
            n,u,p,tel,email=[x.get().strip() for x in fields]
            if not all((n,u,p)): messagebox.showwarning('Datos','Nombre, usuario y contraseña son obligatorios.'); return
            try: execute('INSERT INTO usuarios(nombre,usuario,password_hash,rol,telefono,correo) VALUES(%s,%s,%s,%s,%s,%s)',(n,u,hash_pw(p),role.get(),tel,email)); win.destroy(); self.users()
            except Exception as e: messagebox.showerror('Usuario',str(e))
        self.btn(win,'Crear cuenta',save,True).pack(pady=12)
    def change_password(self):
        uid=self.selected_id('user_table')
        if uid:
            pw=simpledialog.askstring('Contraseña','Nueva contraseña:',show='*',parent=self)
            if pw: execute('UPDATE usuarios SET password_hash=%s WHERE id=%s',(hash_pw(pw),uid)); messagebox.showinfo('Cuenta','Contraseña actualizada.')
    def deactivate_user(self):
        uid=self.selected_id('user_table')
        if uid and uid!=self.user['id'] and messagebox.askyesno('Confirmar','¿Desactivar esta cuenta?'): execute('UPDATE usuarios SET activo=0 WHERE id=%s',(uid,)); self.users()
    def tutor_view(self):
        self.heading('Consulta de información escolar','Vista de consulta de alumnos vinculados a esta cuenta de tutor.')
        rows=query('SELECT a.nombre alumno,g.nombre grupo,a.numero_lista FROM alumnos a JOIN grupos g ON g.id=a.grupo_id WHERE a.tutor_usuario_id=%s AND a.activo=1 ORDER BY a.nombre',(self.user['id'],))
        t=self.table(self.content,['Alumno','Grupo','Número de lista'],[300,160,130],15)
        for r in rows:t.insert('','end',values=(r['alumno'],r['grupo'],r['numero_lista']))

if __name__=='__main__':
    try:
        initialize_db(); App().mainloop()
    except Exception as err:
        try:
            root=tk.Tk(); root.withdraw(); messagebox.showerror('Control Escolar - Inicio',str(err)); root.destroy()
        except Exception: print('Error al iniciar:',err,file=sys.stderr)
