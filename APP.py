from flask import Flask, request, jsonify
from flask_cors import CORS
import subprocess
import os

app = Flask(__name__)
CORS(app)  # Enable Cross-Origin Requests from browser

SCRIPT_DIR = os.path.join(os.getcwd(), "Scripts")  # Path to your scripts folder

@app.route("/run", methods=["POST"])
def run_script():
    data = request.get_json()
    script_name = data.get("script")
    script_path = os.path.join(SCRIPT_DIR, script_name)

    if not os.path.isfile(script_path):
        return jsonify({"error": f"Script not found: {script_name}"}), 404

    try:
        result = subprocess.run(
            ["python", script_path],
            capture_output=True,
            text=True,
            shell=True
        )
        return jsonify({"output": result.stdout})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(port=5000)
