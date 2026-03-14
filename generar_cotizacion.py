#!/usr/bin/env python3
import sys, json, copy, subprocess, shutil, os, tempfile
from docx import Document
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def all_text(el):
    return ''.join((t.text or '') for t in el.iter(f'{{{W}}}t'))

def clear_runs(cell):
    for p in cell.paragraphs:
        for r in p.runs: r.text = ''

def set_cell(cell, texto):
    for p in cell.paragraphs:
        runs = p.runs
        if runs:
            runs[0].text = texto
            for r in runs[1:]: r.text = ''
            return
    cell.paragraphs[0].add_run(texto)

def copiar_fila_despues(tabla, idx):
    tr = tabla.rows[idx]._tr
    tr_nuevo = copy.deepcopy(tr)
    tr.addnext(tr_nuevo)

def generar(datos):
    doc = Document('/app/Cotizacion Qurakuna-PLANTILLA.docx')
    nombre   = datos['nombre']
    productos= datos.get('productos', [])
    traslado = datos.get('traslado', {})
    inc_tras = traslado.get('incluye', False)
    p_tras   = float(traslado.get('precio', 0))
    d_tras   = traslado.get('descripcion', 'Recojo de Plantas & Traslado de Plantas, Personal, Macetas. Ida y Vuelta.')
    pack     = datos.get('pack_cuidado', False)
    base     = datos.get('base_movil', False)

    # 1. Nombre en párrafos
    for p in doc.paragraphs:
        for r in p.runs:
            if '[NOMBRE DEL CLIENTE]' in (r.text or ''):
                r.text = r.text.replace('[NOMBRE DEL CLIENTE]', nombre)

    # 2. Tabla de productos
    t0 = doc.tables[0]
    n_prod = len(productos)
    while len(t0.rows) - 1 < max(n_prod, 1):
        copiar_fila_despues(t0, len(t0.rows) - 1)
    for i, prod in enumerate(productos):
        fila   = t0.rows[i + 1]
        precio = float(prod.get('precio', 0))
        cant   = int(prod.get('cantidad', 1))
        set_cell(fila.cells[0], prod.get('nombre', ''))
        set_cell(fila.cells[1], prod.get('descripcion', ''))
        set_cell(fila.cells[2], f'S/.{precio:,.2f}')
        set_cell(fila.cells[3], str(cant).zfill(2))
        set_cell(fila.cells[4], f'S/.{precio*cant:,.2f}')
    # limpiar filas sobrantes
    for i in range(n_prod + 1, len(t0.rows)):
        for cell in t0.rows[i].cells: clear_runs(cell)

    # 3. Total
    total = sum(float(p.get('precio',0)) * int(p.get('cantidad',1)) for p in productos)
    if inc_tras: total += p_tras
    if pack:     total += 60
    if base:     total += 60

    # 4. Tabla traslado
    t1 = doc.tables[1]
    body_ch = list(doc.element.body)
    if inc_tras and p_tras > 0:
        fila_t = t1.rows[1]
        # descripción
        p_d = fila_t.cells[1].paragraphs[0]
        for r in p_d.runs: r.text = ''
        if p_d.runs: p_d.runs[0].text = d_tras
        else: p_d.add_run(d_tras)
        # precio y costo (runs partidos en p[1], 3 runs: 'S/.' | '00' | '.00')
        for col in [2, 4]:
            # buscar el párrafo que tiene el precio (puede ser p[0] o p[1])
            target_p = None
            for p_c in fila_t.cells[col].paragraphs:
                if p_c.runs:
                    target_p = p_c
                    break
            if target_p is None:
                target_p = fila_t.cells[col].paragraphs[0]
            txt = f'S/. {p_tras:,.2f}'
            runs = target_p.runs
            if runs:
                runs[0].text = txt
                for r in runs[1:]: r.text = ''
            else:
                target_p.add_run(txt)
    else:
        # vaciar título "Servicio Adicional: Traslado"
        for child in body_ch:
            if 'Servicio Adicional' in all_text(child):
                for t_el in child.iter(f'{{{W}}}t'): t_el.text = ''
        for fila in t1.rows:
            for cell in fila.cells: clear_runs(cell)

    # 5. Párrafo del Total (el párrafo vacío entre tabla traslado y "*Precios")
    found_t1 = False
    XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'
    for child in body_ch:
        tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
        if tag == 'tbl' and 'Servicio' in all_text(child):
            found_t1 = True
            continue
        if found_t1 and tag == 'p':
            if '*Precios' in all_text(child) or 'Precios no' in all_text(child):
                break
            # Este es el párrafo del Total
            runs_el = child.findall(f'.//{{{W}}}r')
            total_txt = f'Total: S/.{total:,.2f}'
            # Limpiar todos los <w:t> del párrafo
            for t_el in child.iter(f'{{{W}}}t'): t_el.text = ''
            if runs_el:
                t_els = runs_el[0].findall(f'{{{W}}}t')
                if t_els:
                    t_els[0].text = total_txt
                    t_els[0].set(XML_SPACE, 'preserve')
                else:
                    new_t = etree.SubElement(runs_el[0], f'{{{W}}}t')
                    new_t.text = total_txt
                    new_t.set(XML_SPACE, 'preserve')
            break

    # 6. Opcionales
    t2 = doc.tables[2]
    if not pack and not base:
        # Eliminar completamente todo entre '*Precios no incluyen IGV' y 'Cronograma'
        W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        def all_t(el):
            return ''.join((t.text or '') for t in el.iter('{'+W_NS+'}t'))
        to_remove = []
        in_op = False
        for child in list(doc.element.body):
            txt = all_t(child)
            if 'Precios no incluyen IGV' in txt:
                in_op = True
                continue
            if 'Cronograma' in txt:
                break
            if in_op:
                to_remove.append(child)
        for child in to_remove:
            child.getparent().remove(child)
        # Agregar salto de página después de *Precios no incluyen IGV
        # para que Cronograma siga en página 2
        from docx.oxml import OxmlElement
        # Buscar el párrafo de *Precios y agregarle un page break
        for child in list(doc.element.body):
            txt = all_t(child)
            if 'Precios no incluyen IGV' in txt:
                # Agregar <w:p><w:r><w:br w:type="page"/></w:r></w:p> después
                W_NS2 = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                p_break = etree.SubElement(doc.element.body, f'{{{W_NS2}}}p')
                r_break = etree.SubElement(p_break, f'{{{W_NS2}}}r')
                br = etree.SubElement(r_break, f'{{{W_NS2}}}br')
                br.set(f'{{{W_NS2}}}type', 'page')
                # Mover el salto justo después del párrafo de *Precios
                child.addnext(p_break)
                break
    else:
        if not pack:
            for cell in t2.rows[1].cells: clear_runs(cell)
        if not base:
            for cell in t2.rows[2].cells: clear_runs(cell)

    # 7. Guardar y convertir
    tmpdir = tempfile.mkdtemp()
    safe   = nombre.replace(' ','_').replace('/','_')
    docx_p = os.path.join(tmpdir, f'Cotizacion_Qurakuna_{safe}.docx')
    doc.save(docx_p)

    subprocess.run(
        ['libreoffice','--headless','--convert-to','pdf','--outdir',tmpdir,docx_p],
        capture_output=True, text=True, timeout=60
    )

    pdf_p = docx_p.replace('.docx', '.pdf')
    os.makedirs('/tmp/cotizaciones', exist_ok=True)
    out   = f'/tmp/cotizaciones/Cotizacion_Qurakuna_{safe}.pdf'
    if os.path.exists(pdf_p):
        shutil.copy(pdf_p, out)
        print(out)
    else:
        out_d = out.replace('.pdf','.docx')
        shutil.copy(docx_p, out_d)
        print(out_d)

    shutil.rmtree(tmpdir, ignore_errors=True)

if __name__ == '__main__':
    datos = json.loads(sys.argv[1])
    generar(datos)
