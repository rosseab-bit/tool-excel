import threading

import pytest

from excel_tool.core.errors import Cancelado, ExcelToolError
from excel_tool.core.progress import CancelToken, ProgressEvent, ThrottledProgress


def test_cancel_token_basico():
    token = CancelToken()
    assert not token.cancelled
    token.check()
    token.cancel()
    assert token.cancelled
    with pytest.raises(Cancelado) as exc:
        token.check()
    assert isinstance(exc.value, ExcelToolError)
    assert exc.value.mensaje == "Operación cancelada."


def test_cancel_token_desde_otro_hilo():
    token = CancelToken()
    t = threading.Thread(target=token.cancel)
    t.start()
    t.join()
    assert token.cancelled
    with pytest.raises(Cancelado):
        token.check()


def test_cancel_token_corta_un_hilo_de_trabajo():
    token = CancelToken()
    result = []

    def work():
        try:
            while True:
                token.check()
        except Cancelado:
            result.append("cancelado")

    t = threading.Thread(target=work)
    t.start()
    token.cancel()
    t.join(timeout=2)
    assert result == ["cancelado"]


def test_throttle_sin_callback():
    ThrottledProgress(None)(ProgressEvent("x", 1, 10))


def test_throttle_limita_frecuencia():
    now = [0.0]
    got = []
    report = ThrottledProgress(got.append, clock=lambda: now[0])
    for i in range(100):
        report(ProgressEvent("x", i, 1000))
        now[0] += 0.01  # 100 eventos por segundo
    assert 9 <= len(got) <= 11
    assert got[0].done == 0


def test_throttle_deja_pasar_el_ultimo_y_force():
    now = [0.0]
    got = []
    report = ThrottledProgress(got.append, clock=lambda: now[0])
    report(ProgressEvent("x", 1, 10))
    report(ProgressEvent("x", 2, 10))  # descartado
    report(ProgressEvent("x", 10, 10))  # último: pasa
    report.force(ProgressEvent("y", 0, None))
    assert [e.done for e in got] == [1, 10, 0]
