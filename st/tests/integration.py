import os
import sys
import pytest
import tempfile
import pandas as pd
from openpyxl import Workbook
from api import app

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# Define paths
ROOT_DIR = os.getcwd()
UPLOADS_DIR = os.path.join(ROOT_DIR, "uploaded_from_main_uploads")
EXISTING_EXCEL_FILE = os.path.join(UPLOADS_DIR, "test_input.xlsx")

# Create a Flask test client
@pytest.fixture
def client():
    app.config['TESTING'] = True
    with app.test_client() as client:
        yield client

# Fixture to set up the test file before each test and remove it after
@pytest.fixture(autouse=True)
def setup_and_teardown():
    # Ensure uploads directory exists
    if not os.path.exists(UPLOADS_DIR):
        os.makedirs(UPLOADS_DIR)

    # Create an existing test_input.xlsx file with required data if it doesn't exist
    if not os.path.exists(EXISTING_EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        headers = ["TestCaseID", "Method", "URL", "Headers", "Payload", "ExpectedStatusCode", "ExpectedResponse"]
        ws.append(headers)
        data = [
            (1, 'GET', 'https://jsonplaceholder.typicode.com/posts/1', '', '', 200, 'sunt aut facere repellat provident occaecati excepturi optio reprehenderit'),
            (2, 'POST', 'https://jsonplaceholder.typicode.com/posts', '{"Content-type": "application/json; charset=UTF-8"}', '{"title": "foo", "body": "bar", "userId": 2}', 201, '{"title": "foo", "body": "bar", "userId": 1, "id": 101}'),
            (3, 'PUT', 'https://jsonplaceholder.typicode.com/posts/1', '{"Content-type": "application/json; charset=UTF-8"}', '{"id": 1, "title": "foo", "body": "bar", "userId": 1}', 200, '{"id": 1, "title": "foo", "body": "bar", "userId": 1}'),
            (4, 'DELETE', 'https://jsonplaceholder.typicode.com/posts/1', '', '', 200, '')
        ]
        for row in data:
            ws.append(row)
        wb.save(EXISTING_EXCEL_FILE)

    # Yield the path of the existing Excel file for use in tests
    yield EXISTING_EXCEL_FILE

def test_home_page(client):
    response = client.get('/')
    assert response.status_code == 302

def test_upload_file(client, setup_and_teardown):
    existing_file = setup_and_teardown

    # Simulate file upload
    data = {
        'file': (open(existing_file, 'rb'), 'test_input.xlsx')
    }
    response = client.post('/upload', data=data, content_type='multipart/form-data')
    assert response.status_code == 302  # Redirects to run_tests

def test_run_tests(client, setup_and_teardown):
    existing_file = setup_and_teardown

    # Simulate running tests
    filename = os.path.basename(existing_file)
    response = client.get(f'/run_tests/{filename}')
    assert response.status_code == 200
    assert b'Tests completed.' in response.data
