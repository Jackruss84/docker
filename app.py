"""
Valoria PDF Service — Micro-API Flask
Génère l'avis de valeur propriétaire en PDF
"""
import os
import io
import json
from flask import Flask, request, send_file, jsonify
from pdf_generator import generate_owner_pdf

app = Flask(__name__)

# Clé secrète partagée avec la Edge Function Supabase
PDF_SECRET = os.environ.get('PDF_SECRET', '')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'valoria-pdf'})

@app.route('/generate', methods=['POST'])
def generate():
    # Auth check
    auth = request.headers.get('X-PDF-Secret', '')
    if PDF_SECRET and auth != PDF_SECRET:
        return jsonify({'error': 'Unauthorized'}), 401

    try:
        body = request.get_json()
        if not body:
            return jsonify({'error': 'Body JSON requis'}), 400

        report_data = body.get('report', {})
        agency_data = body.get('agency', {})

        if not report_data:
            return jsonify({'error': 'report requis'}), 400

        # Générer dans un buffer mémoire
        buf = io.BytesIO()
        generate_owner_pdf(report_data, agency_data, buf)
        buf.seek(0)

        address = report_data.get('address', 'avis-valeur')
        safe_name = ''.join(c if c.isalnum() or c in '-_ ' else '-' for c in address)
        safe_name = safe_name.lower().replace(' ', '-')[:50]
        filename = f"{safe_name}-avis-valeur.pdf"

        return send_file(
            buf,
            mimetype='application/pdf',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        print(f'[PDF] Error: {e}')
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
