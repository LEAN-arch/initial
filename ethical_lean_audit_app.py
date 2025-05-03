from flask import Flask, render_template, request, send_file
import pandas as pd
import json
import io
import base64
import datetime
from collections import defaultdict

app = Flask(__name__)

# Questions data
questions = {
    "Employee Experience": {
        "English": [
            "Does work contribute to organizational success and personal growth?",
            "Are leaders evaluated on creating positive work environments?"
        ],
        "Español": [
            "¿Contribuye el trabajo al éxito organizacional y crecimiento personal?",
            "¿Se evalúa a los líderes por crear entornos positivos?"
        ]
    },
    "Process Improvement": {
        "English": [
            "Are human impacts considered when eliminating waste?",
            "Are process changes evaluated for negative consequences?"
        ],
        "Español": [
            "¿Se consideran impactos humanos al eliminar desperdicios?",
            "¿Se evalúan los cambios por consecuencias negativas?"
        ]
    }
}

# Store responses in memory
responses = defaultdict(lambda: [0] * 2)
completed_categories = set()
current_category = 0
lang = "English"
timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

@app.route("/", methods=["GET", "POST"])
def index():
    global responses, completed_categories, current_category, lang
    categories = list(questions.keys())
    
    if request.method == "POST":
        action = request.form.get("action")
        if action == "submit":
            category = request.form.get("category")
            scores = [int(request.form.get(f"q{idx}", 0)) for idx in range(len(questions[category][lang]))]
            responses[category] = scores
            if all(1 <= score <= 5 for score in scores):
                completed_categories.add(category)
        elif action == "next":
            current_category = min(current_category + 1, len(categories) - 1)
        elif action == "prev":
            current_category = max(current_category - 1, 0)
        elif action == "lang":
            lang = request.form.get("lang", "English")
        elif action == "reset":
            responses.clear()
            completed_categories.clear()
            current_category = 0
            lang = "English"
        elif action == "report":
            incomplete = [cat for cat in categories if not all(1 <= score <= 5 for score in responses.get(cat, [0] * 2))]
            if not incomplete:
                results = []
                for cat in categories:
                    scores = responses.get(cat, [0] * 2)
                    total = sum(scores)
                    percent = (total / (len(scores) * 5)) * 100
                    results.append({"Category": cat, "Score": total, "Percent": percent})
                df = pd.DataFrame(results)
                csv_buffer = io.StringIO()
                csv_buffer.write(f"Ethical Lean Audit Report\nDate: {timestamp}\n\n")
                df.to_csv(csv_buffer, index=False)
                csv_data = csv_buffer.getvalue()
                b64_csv = base64.b64encode(csv_data.encode()).decode()
                return json.dumps({"csv": b64_csv})
    
    category = categories[current_category]
    score_sum = sum(responses[category])
    max_score = len(questions[category][lang]) * 5
    score_percent = (score_sum / max_score) * 100 if max_score > 0 else 0
    completed_count = len(completed_categories)
    
    return render_template("index.html", 
                         questions=questions[category][lang],
                         category=category,
                         categories=categories,
                         current_category=current_category,
                         lang=lang,
                         labels=["Not Practiced", "Rarely", "Partially", "Mostly", "Fully"] if lang == "English" else ["No practicado", "Raramente", "Parcialmente", "Mayormente", "Totalmente"],
                         scores=responses[category],
                         score_sum=score_sum,
                         max_score=max_score,
                         score_percent=score_percent,
                         completed_count=completed_count,
                         total_categories=len(categories),
                         completed_categories=completed_categories)

@app.route("/download")
def download():
    categories = list(questions.keys())
    results = []
    for cat in categories:
        scores = responses.get(cat, [0] * 2)
        total = sum(scores)
        percent = (total / (len(scores) * 5)) * 100
        results.append({"Category": cat, "Score": total, "Percent": percent})
    df = pd.DataFrame(results)
    csv_buffer = io.StringIO()
    csv_buffer.write(f"Ethical Lean Audit Report\nDate: {timestamp}\n\n")
    df.to_csv(csv_buffer, index=False)
    return send_file(
        io.BytesIO(csv_buffer.getvalue().encode()),
        mimetype='text/csv',
        as_attachment=True,
        download_name="audit_results.csv"
    )

if __name__ == "__main__":
    app.run(debug=True)
