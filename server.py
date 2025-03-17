from cvgen import cv_generator
import json
import traceback
from flask import Flask, request, jsonify
import threading

app = Flask(__name__)
lock = threading.Lock()
is_generating = False


@app.route('/generate_cvs', methods=['POST'])
def generate_cvs():
    global is_generating
    if lock.locked():
        return jsonify({"id": None, "message": "CV generation already in progress. Try again later"}), 429
    with lock:
        is_generating = True
        try:
            payload_data = json.loads(request.stream.read().decode())
            cv_record_id = payload_data['cv_record_id']
            results = cv_generator(cv_record_id)
        except Exception as e:
            print(e)
            print(traceback.format_exc())
            results = {}
        finally:
            is_generating = False
    return jsonify(results)


app.run(host='0.0.0.0', port=7621)

