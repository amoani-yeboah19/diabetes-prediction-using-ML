import unittest
import warnings

from sklearn.exceptions import InconsistentVersionWarning

from app.app import FEATURE_NAMES, app, load_artifacts


class PredictionTests(unittest.TestCase):
    def test_artifacts_match_runtime(self):
        load_artifacts.cache_clear()
        with warnings.catch_warnings():
            warnings.simplefilter('error', InconsistentVersionWarning)
            load_artifacts()

    def test_pages_and_stylesheet(self):
        with app.test_client() as client:
            for route in ('/', '/predict', '/style.css'):
                with self.subTest(route=route):
                    with client.get(route) as response:
                        self.assertEqual(response.status_code, 200)
            with client.get('/style.css') as response:
                self.assertIn('text/css', response.content_type)

    def test_invalid_input_is_rejected(self):
        values = dict(zip(FEATURE_NAMES, [1, 85, 66, 29, 94, 26.6, 0.351, 31]))
        with app.test_client() as client:
            for invalid in ('', 'abc', 'nan', 'inf'):
                with self.subTest(invalid=invalid):
                    response = client.post('/predict', data={**values, 'Glucose': invalid})
                    self.assertEqual(response.status_code, 200)
                    self.assertIn('valid, finite numbers', response.get_data(as_text=True))
                    self.assertNotIn('Diabetes probability:', response.get_data(as_text=True))
            response = client.post('/predict', data={**values, 'model': 'Unknown'})
            self.assertNotIn('Diabetes probability:', response.get_data(as_text=True))

    def test_saved_models_predict_through_form(self):
        app.config['TESTING'] = True
        cases = [
            ([1, 85, 66, 29, 94, 26.6, 0.351, 31], 'Not Diabetic'),
            ([8, 183, 64, 33, 180, 23.3, 0.672, 32], 'Diabetic'),
        ]
        with app.test_client() as client:
            for model in load_artifacts()[1]:
                for values, expected in cases:
                    with self.subTest(model=model, values=values):
                        data = dict(zip(FEATURE_NAMES, values))
                        data['model'] = model
                        response = client.post('/predict', data=data)
                        self.assertEqual(response.status_code, 200)
                        html = response.get_data(as_text=True)
                        self.assertIn('Diabetes probability:', html)
                        self.assertIn(f'<em>{model}</em>', html)
                        if model == 'Logistic Regression':
                            self.assertIn(f'class="result-label">{expected}</p>', html)


if __name__ == '__main__':
    unittest.main()
