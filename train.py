import os
import joblib
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, precision_score, recall_score, f1_score

from preprocess import preprocess_pipeline

DATA_PATH = os.path.join('data', 'diabetes.csv')
MODELS_DIR = 'models'
SCALER_PATH = os.path.join(MODELS_DIR, 'scaler.pkl')


def get_models():
    return {
        'Logistic Regression': LogisticRegression(max_iter=1000, random_state=42),
        'Decision Tree': DecisionTreeClassifier(random_state=42),
        'Random Forest': RandomForestClassifier(n_estimators=100, random_state=42),
    }


def evaluate(model, X, y, split_name):
    y_pred = model.predict(X)
    acc = accuracy_score(y, y_pred)
    prec = precision_score(y, y_pred, zero_division=0)
    rec = recall_score(y, y_pred, zero_division=0)
    f1 = f1_score(y, y_pred, zero_division=0)
    print(f"  [{split_name}] Accuracy: {acc:.4f} | Precision: {prec:.4f} | Recall: {rec:.4f} | F1: {f1:.4f}")
    return {'accuracy': acc, 'precision': prec, 'recall': rec, 'f1': f1}


def train_and_evaluate():
    print("Loading and preprocessing data...")
    X_train, X_val, X_test, y_train, y_val, y_test = preprocess_pipeline(
        DATA_PATH, scaler_path=SCALER_PATH
    )[:6]

    os.makedirs(MODELS_DIR, exist_ok=True)
    models = get_models()
    results = {}

    for name, model in models.items():
        print(f"\n--- {name} ---")
        model.fit(X_train, y_train)

        train_metrics = evaluate(model, X_train, y_train, 'Train')
        val_metrics = evaluate(model, X_val, y_val, 'Val')
        test_metrics = evaluate(model, X_test, y_test, 'Test')

        results[name] = {
            'train': train_metrics,
            'val': val_metrics,
            'test': test_metrics,
        }

        model_file = os.path.join(MODELS_DIR, f"{name.lower().replace(' ', '_')}.pkl")
        joblib.dump(model, model_file)
        print(f"  Saved to {model_file}")

    print("\n=== FINAL TEST SET COMPARISON ===")
    best_name, best_acc = None, 0
    for name, res in results.items():
        acc = res['test']['accuracy']
        print(f"  {name}: Accuracy={acc:.4f}, Precision={res['test']['precision']:.4f}, Recall={res['test']['recall']:.4f}")
        if acc > best_acc:
            best_acc, best_name = acc, name

    print(f"\nBest model: {best_name} (Test Accuracy: {best_acc:.4f})")
    return results


if __name__ == '__main__':
    train_and_evaluate()
