import os
from functools import lru_cache
import numpy as np
import joblib
from flask import Flask, render_template, request, send_from_directory

app = Flask(__name__)

MODELS_DIR = os.path.join(os.path.dirname(__file__), '..', 'models')

FEATURE_NAMES = [
    'Pregnancies', 'Glucose', 'BloodPressure', 'SkinThickness',
    'Insulin', 'BMI', 'DiabetesPedigreeFunction', 'Age'
]

@lru_cache(maxsize=1)
def load_artifacts():
    """Load the trained artefacts once per running serverless instance."""
    scaler = joblib.load(os.path.join(MODELS_DIR, 'scaler.pkl'))
    models = {
        'Logistic Regression': joblib.load(os.path.join(MODELS_DIR, 'logistic_regression.pkl')),
        'Decision Tree': joblib.load(os.path.join(MODELS_DIR, 'decision_tree.pkl')),
        'Random Forest': joblib.load(os.path.join(MODELS_DIR, 'random_forest.pkl')),
    }
    return scaler, models


@app.route('/style.css')
def stylesheet():
    return send_from_directory(os.path.join(app.root_path, '..', 'public'), 'style.css')


@app.route('/')
def landing():
    """Project introduction shown before the prediction tool."""
    return render_template('landing.html')


@app.route('/predict', methods=['GET', 'POST'])
def predict():
    prediction = None
    probability = None
    selected_model = None
    form_values = {}
    error = None

    scaler, models = load_artifacts()

    if request.method == 'POST':
        try:
            selected_model = request.form.get('model', 'Logistic Regression')
            if selected_model not in models:
                raise ValueError('Please select a valid prediction model.')

            form_values = {f: request.form.get(f, '') for f in FEATURE_NAMES}

            features = np.array([[float(form_values[f]) for f in FEATURE_NAMES]])
            if not np.isfinite(features).all():
                raise ValueError('Please enter finite numeric values in every field.')

            # Preserve the feature names expected by the saved scaler without
            # requiring pandas in the web runtime.
            if hasattr(scaler, 'feature_names_in_'):
                features_scaled = (features - scaler.mean_) / scaler.scale_
            else:
                features_scaled = scaler.transform(features)

            model = models[selected_model]
            pred = model.predict(features_scaled)[0]
            prob = model.predict_proba(features_scaled)[0][1] * 100

            prediction = 'Diabetic' if pred == 1 else 'Not Diabetic'
            probability = round(prob, 1)

        except (TypeError, ValueError):
            error = "Please fill in all fields with valid, finite numbers."

    return render_template(
        'index.html',
        prediction=prediction,
        probability=probability,
        selected_model=selected_model,
        model_names=list(models.keys()),
        feature_names=FEATURE_NAMES,
        form_values=form_values,
        error=error,
    )


if __name__ == '__main__':
    app.run()
