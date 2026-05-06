from fantozzi.normalization import normalize


def test_agua_sanitaria():
    assert normalize("ÁGUA SANITÁRIA") == "agua sanitaria"


def test_coca_cola_lata():
    assert normalize("COCA-COLA LATA") == "coca cola lata"


def test_rio_sol():
    assert normalize("Rio Sol") == "rio sol"


def test_empty_string():
    assert normalize("") == ""
