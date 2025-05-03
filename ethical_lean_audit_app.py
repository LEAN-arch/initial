import pandas as pd
import io
import datetime
import sys

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

# Likert scale labels
labels = {
    "English": ["Not Practiced", "Rarely", "Partially", "Mostly", "Fully"],
    "Español": ["No practicado", "Raramente", "Parcialmente", "Mayormente", "Totalmente"]
}

def main():
    categories = list(questions.keys())
    responses = {cat: [0, 0] for cat in categories}
    completed_categories = set()
    lang = input("Select language (English/Español): ").strip().capitalize()
    if lang not in ["English", "Español"]:
        lang = "English"
    current_category = 0
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    while True:
        category = categories[current_category]
        print(f"\n=== {category} ===")
        scores = responses[category]
        for idx, q in enumerate(questions[category][lang]):
            print(f"\n{q}")
            print("Rate 1-5 (1: {}, 5: {}):".format(labels[lang][0], labels[lang][4]))
            while True:
                try:
                    score = int(input("Enter score (1-5, 0 to skip): "))
                    if score == 0:
                        break
                    if 1 <= score <= 5:
                        scores[idx] = score
                        break
                    print("Invalid score. Enter 1-5 or 0 to skip.")
                except ValueError:
                    print("Invalid input. Enter a number 1-5 or 0 to skip.")

        if all(1 <= score <= 5 for score in scores):
            completed_categories.add(category)

        score_sum = sum(scores)
        max_score = len(questions[category][lang]) * 5
        score_percent = (score_sum / max_score) * 100 if max_score > 0 else 0
        print(f"\nScore: {score_sum}/{max_score} ({score_percent:.1f}%)")
        if category in completed_categories:
            print(f"{'Category completed!' if lang == 'English' else '¡Categoría completada!'}")
        print(f"Progress: {len(completed_categories)}/{len(categories)} completed ({(len(completed_categories)/len(categories))*100:.1f}%)")

        print("\nOptions: (p)revious, (n)ext, (r)eset, (g)enerate report, (q)uit")
        choice = input("Choose an option: ").strip().lower()
        if choice == "p" and current_category > 0:
            current_category -= 1
        elif choice == "n" and current_category < len(categories) - 1:
            current_category += 1
        elif choice == "r":
            responses = {cat: [0, 0] for cat in categories}
            completed_categories.clear()
            current_category = 0
            lang = input("Select language (English/Español): ").strip().capitalize()
            if lang not in ["English", "Español"]:
                lang = "English"
        elif choice == "g":
            incomplete = [cat for cat in categories if not all(1 <= score <= 5 for score in responses.get(cat, [0, 0]))]
            if incomplete:
                print(f"{'Complete all questions for: ' if lang == 'English' else 'Complete todas las preguntas para: '}{', '.join(incomplete)}")
            else:
                results = []
                for cat in categories:
                    scores = responses.get(cat, [0, 0])
                    total = sum(scores)
                    percent = (total / (len(scores) * 5)) * 100
                    results.append({"Category": cat, "Score": total, "Percent": percent})
                df = pd.DataFrame(results)
                csv_buffer = io.StringIO()
                csv_buffer.write(f"Ethical Lean Audit Report\nDate: {timestamp}\n\n")
                df.to_csv(csv_buffer, index=False)
                with open("audit_results.csv", "w") as f:
                    f.write(csv_buffer.getvalue())
                print(f"{'Report saved as audit_results.csv' if lang == 'English' else 'Informe guardado como audit_results.csv'}")
        elif choice == "q":
            break

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        with open("error.log", "w") as f:
            f.write(f"Error: {str(e)}")
        sys.exit(f"Failed to start app: {str(e)}")
