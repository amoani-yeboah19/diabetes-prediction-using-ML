# Diabetes Prediction System

A Flask web application that predicts diabetes risk from eight health measurements using pre-trained machine-learning models.

## Run locally

```bash
python3.12 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m flask --app app.app run
```

Open `http://127.0.0.1:5000` in a browser. This project supports Python 3.11 or 3.12; its pinned dependencies are not compatible with Python 3.14. To retrain the models, install `requirements-dev.txt` and run `python train.py` from the activated virtual environment.

The saved models require scikit-learn 1.8.0. After updating an existing installation, run `python -m pip install -r requirements.txt` and restart Flask. Keep the scikit-learn version in `requirements.txt` and `pyproject.toml` aligned with the version used to train the models; loading these artifacts with 1.5.0 can break logistic regression probability prediction.

Run `python -m unittest discover -s tests -v` to check artifact compatibility and web-form predictions for all three models.

## Deploy to Vercel

This repository is configured for Vercel's Python/Flask runtime. The pre-trained files in `models/` are deployed with the application; no database or environment variables are required.

1. Import [this GitHub repository](https://github.com/amoani-yeboah19/diabetes-prediction-using-ML) into Vercel.
2. At [vercel.com/new](https://vercel.com/new), import that GitHub repository.
3. Keep the project root as this folder and click **Deploy**. Do not set a Build Command or Output Directory.
4. Vercel will give you a public `*.vercel.app` link. Every push to the selected branch will deploy an update.

Alternatively, after installing and logging into the Vercel CLI, run `vercel` in this folder and follow the prompts.

`public/style.css` is served through Vercel's CDN. The Flask application is explicitly declared in `pyproject.toml`, and the model files are loaded once for each warm serverless instance.

## Project files

- `app/`: Flask routes and HTML templates.
- `public/style.css`: shared stylesheet for local Flask and Vercel.
- `data/diabetes.csv`: dataset used for training and exploration.
- `preprocess.py` and `train.py`: preprocessing, model training, and evaluation.
- `models/`: saved scaler and three classifiers required to run the application.
- `notebooks/01_eda.ipynb`: exploratory data analysis.
- `tests/`: model compatibility and web application checks.
- `generate_report.js`: optional Word report generator; Node.js is needed only for this utility.

On Windows with Python 3.12 installed, run `setup_and_run.bat` to create an environment, install development dependencies, train the models, and start Flask. Training replaces the saved model files.

## Optional report generation

```bash
npm ci
npm run report
```

This creates `Diabetes_Project_Report.docx` in the project folder. Use `npm run report -- /path/to/report.docx` for a different output path. The report contains static narrative, metrics, and source-code excerpts; review them against the current training output before submitting it.

## Evaluation limitations

The dataset is identified in the EDA notebook as the Pima Indians Diabetes Dataset and is not representative of every population. The current preprocessing replaces missing values using medians calculated before the train/validation/test split. This can leak information between splits; reported metrics should be treated as preliminary. A future evaluation should fit imputation on the training split only and assess probability calibration.

## Important note

This tool is an academic demonstration, not a clinical diagnostic device. Its result should never substitute for medical advice or professional diagnosis.
