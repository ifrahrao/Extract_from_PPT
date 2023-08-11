import sys
from helpers import pptToJson

from flask import Flask, request,jsonify
from flask_cors import CORS
app = Flask(__name__)
CORS(app)

@app.route('/extract_from_ppt', methods=['POST'])
def execute_python():
    pdf_file_path = request.json['file_path']
    print(f"PDF file path: {pdf_file_path}")


    result=pptToJson(pdf_file_path)
    return jsonify({'result': result})

if __name__ == '__main__':
    app.run()



