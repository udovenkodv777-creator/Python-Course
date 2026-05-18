import pytest

def add(a, b):
    return a + b

def divide(a, b):
    if b == 0:
        raise ValueError("Division by zero")
    return a / b

@pytest.mark.parametrize("a,b,expected", [
    (1, 2, 3),
    (0, 0, 0),
    (-1, 1, 0),
])
def test_add(a, b, expected):
    assert add(a, b) == expected
    
def test_divide_by_zero():
    with pytest.raises(ValueError):
        divide(10, 0)  
          
def test_divide():
    assert divide(10, 2) ==5
         
@pytest.fixture
def sample_data():
    return {"a": 10, "b": 5}

def test_with_fixture(sample_data):
    assert divide(sample_data["a"], sample_data["b"]) == 2