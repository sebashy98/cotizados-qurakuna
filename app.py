import os, json, tempfile, shutil, subprocess, sys
from flask import Flask, request, send_file, render_template, jsonify

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generar', methods=['POST'])
def generar():
    try:
        datos = request.get_json()
        if not datos or not datos.get('nombre'):
            return 'Falta el nombre del cliente', 400

        # Llamar al script generador
        result = subprocess.run(
            [sys.executable, '/app/generar_cotizacion.py', json.dumps(datos)],
            capture_output=True, text=True, timeout=90
        )

        pdf_path = result.stdout.strip()
        if not pdf_path or not os.path.exists(pdf_path):
            return f'Error generando PDF: {result.stderr[:300]}', 500

        nombre = datos['nombre'].replace(' ', '_').replace('/', '_')
        return send_file(
            pdf_path,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=f'Cotizacion_Qurakuna_{nombre}.pdf'
        )
    except Exception as e:
        return str(e), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8765))
    app.run(host='0.0.0.0', port=port)
