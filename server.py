# server.py
from flask import Flask, request, jsonify
from analysis_runner import run_analysis




app = Flask(__name__)

@app.route('/run-analysis', methods=['POST'])
def handle_run_analysis():
    import logging
    logging.basicConfig(filename=r"C:\Users\pli\Desktop\flask_debug.log", level=logging.DEBUG)


    logging.info("Received request.")
    print("✅ Flask received request")

    data = request.json
    print(f"Received JSON: {data}")
    
    
    try:
        data = request.get_json()
        
        workbook_path = data['workbook_path']
        target_sheet = data['target_sheet']
        target_variable = data['target_variable']
        target_period = data['target_period']
        target_cell = data['target_cell']
        baseline_comparison_cell = data['baseline_comparison_cell']
        print('Start running!!!!!!!!!!!!!')
        result = run_analysis(
            workbook_path=workbook_path,
            target_sheet=target_sheet,
            target_variable=target_variable,
            target_period=target_period,
            target_cell=target_cell,
            baseline_comparison_cell=baseline_comparison_cell
        )

        return jsonify(result)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    print('!!!!!!!!!!!!!!!!!!!!!!!')
    app.run(debug = True, host='0.0.0.0', port=5001)
