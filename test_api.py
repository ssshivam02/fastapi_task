from fastapi.testclient import TestClient
import pytest
# html report pytest --html=html_report\Report.html --tb=no --self-contained-html .\test_api.py
@pytest.fixture
def client():
    from main import app
    with TestClient(app) as client:
        yield client

def test_get_list_employee(client):

    response = client.get(url = "/get_list_employee_details")

    print(response.json())

def test_add_employee(client):
    header_data = {
        "Content-Type" : "application/json"
    }
    body = {
        "E.Code": "100",
        "E.Name" : "Sharma", 
        "Date": "05-08-2023",
        "Attendance": "A"
    }
    reponse = client.post(url = "/add_employee", headers = header_data, json = body)

    print(reponse.json())


def test_get_employee_details(client):
    employee_code = "100"
    response = client.get(url = f"/employee_detail/{employee_code}")
    print(response.json())


def test_get_employee_detail_absent(client):
    response = client.get(url = "/employee_detail/absent/list")
    print(response.json())


def test_get_send_email(client):
    response = client.get(url = "/send_email")
    print(response.json())