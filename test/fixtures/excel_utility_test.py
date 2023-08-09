import pytest
import os

from app import excel_utility as eu


@pytest.fixture
def utility():
    return eu.ExcelUtility()


@pytest.fixture
def save_directory():
    try:
        os.makedirs("test/fixtures/output/test_outputs")
    except:
        pass
    return "test/fixtures/output/test_outputs/"


def test_final_file_cleanup(utility, save_directory):
    utility.filename = "testFile"
    utility.save_directory = save_directory

    with open(save_directory + "_testFile.csv", "w") as f:
        f.write("TEST")

    file_exists = os.path.exists(save_directory + "_testFile.csv")

    assert file_exists

    utility.final_file_clean_up("csv")

    assert not os.path.exists(save_directory + "_testFile.csv")
