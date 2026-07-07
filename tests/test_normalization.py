from fantozzi.normalization import normalize, normalize_cnpj


def test_agua_sanitaria():
    assert normalize("ÁGUA SANITÁRIA") == "agua sanitaria"


def test_coca_cola_lata():
    assert normalize("COCA-COLA LATA") == "coca cola lata"


def test_rio_sol():
    assert normalize("Rio Sol") == "rio sol"


def test_empty_string():
    assert normalize("") == ""


def test_normalize_cnpj_strips_punctuation():
    assert normalize_cnpj("42.870.341/0001-68") == "42870341000168"


def test_normalize_cnpj_already_digits():
    assert normalize_cnpj("55746722000181") == "55746722000181"


def test_normalize_cnpj_empty():
    assert normalize_cnpj("") == ""


def test_normalize_cnpj_none():
    assert normalize_cnpj(None) == ""
