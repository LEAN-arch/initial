import pandas as pd
import io
import datetime
import sys
import os
from typing import Dict, List, Set

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

def validate_score(score: str) -> int:
    """Validate and convert score input to integer."""
    try:
        score_int = int(score)
        if score_int not in range(0, 6):  # 0 to 5 inclusive
            raise ValueError("Score must be 0-5")
        return score_int
    except ValueError as e:
        raise ValueError("Invalid input. Enter a number 0-5") from e

def generate_report(responses: Dict[str, List[int]], categories: List[str], 
                   timestamp: str, lang: str) -> str:
    """Generate and save audit report."""
    results = []
    for cat in categories:
        scores = responses.get(cat, [0, 0])
        total = sum(scores)
        max_score = len(scores) * 5
        percent = (total / max_score * 100) if max_score > 0 else 0
        results.append({
            "Category": cat,
            "Score": total,
            "Max Score": max_score,
            "Percent": round(percent, 1)
        })
    
    df = pd.DataFrame(results)
    csv_buffer = io.StringIO()
    csv_buffer.write(f"Ethical Lean Audit Report\nDate: {timestamp}\nLanguage: {lang}\n\n")
    df.to_csv(csv_buffer, index=False)
    
    try:
        with open("audit_results.csv", "w", encoding='utf-8') as f:
            f.write(csv_buffer.getvalue())
        return f"{'Report saved as audit_results.csv' if lang == 'English' else 'Informe guardado como audit_results.csv'}"
    except IOError as e:
        return f"{'Failed to save report: ' if lang == 'English' else 'Error al guardar informe: '}{str(e)}"

def main():
    categories = list(questions.keys())
    responses: Dict[str, List[int]] = {cat: [0] * len(questions[cat]["English"]) for cat in categories}
    completed_categories: Set[str] = set()
    
    # Initialize language with validation
    while True:
        lang = input("Select language (English/Español): ").strip().capitalize()
        if lang in ["English", "Español"]:
            break
        print("Invalid language. Please select English or Español.")
    
    current_category = 0
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    while True:
        try:
            category = categories[current_category]
            print(f"\n=== {category} ===")
            scores = responses[category]
            
            for idx, q in enumerate(questions[category][lang]):
                while True:
                    print(f"\n{q}")
                    print(f"Rate 1-5 (1: {labels[lang][0]}, 5: {labels[lang][4]}):")
                    try:
                        score = validate_score(input("Enter score (1-5, 0 to skip): "))
                        if score == 0:
                            break
                        scores[idx] = score
                        break
                    except ValueError as e:
                        print(str(e))

            # Update completion status
            if all(1 <= score <= 5 for score in scores):
                completed_categories.add(category)
            else:
                completed_categories.discard(category)

            # Display category score
            score_sum = sum(scores)
            max_score = len(questions[category][lang]) * 5
            score_percent = (score_sum / max_score * 100) if max_score > 0 else 0
            print(f"\nScore: {score_sum}/{max_score} ({score_percent:.1f}%)")
            if category in completed_categories:
                print(f"{'Category completed!' if lang == 'English' else '¡Categoría completada!'}")
            print(f"Progress: {len(completed_categories)}/{len(categories)} completed "
                  f"({(len(completed_categories)/len(categories))*100:.1f}%)")

            # Navigation options
            print("\nOptions: (p)revious, (n)ext, (r)eset, (g)enerate report, (q)uit")
            choice = input("Choose an option: ").strip().lower()
            
            if choice == "p" and current_category > 0:
                current_category -= 1
            elif choice == "n" and current_category < len(categories) - 1:
                current_category += 1
            elif choice == "r":
                responses = {cat: [0] * len(questions[cat]["English"]) for cat in categories}
                completed_categories.clear()
                current_category = 0
                while True:
                    lang = input("Select language (English/Español): ").strip().capitalize()
                    if lang in ["English", "Español"]:
                        break
                    print("Invalid language. Please select English or Español.")
            elif choice == "g":
                incomplete = [cat for cat in categories 
                            if not all(1 <= score <= 5 for score in responses.get(cat, [0, 0]))]
                if incomplete:
                    print(f"{'Complete all questions for: ' if lang == 'English' else 'Complete todas las preguntas para: '}"
                          f"{', '.join(incomplete)}")
                else:
                    print(generate_report(responses, categories, timestamp, lang))
            elif choice == "q":
                break
            else:
                print("Invalid option. Please choose p, n, r, g, or q.")
                
        except KeyboardInterrupt:
            print("\n{'Program interrupted. Exiting...' if lang == 'English' else 'Programa interrumpido. Saliendo...'}")
            break
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            print(error_msg)
            with open("error.log", "a", encoding='utf-8') as f:
                f.write(f"{timestamp}: {error_msg}\n")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        error_msg = f"Fatal error: {str(e)}"
        with open("error.log", "a", encoding='utf-8') as f:
            f.write(f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}: {error_msg}\n")
        sys.exit(error_msg)
