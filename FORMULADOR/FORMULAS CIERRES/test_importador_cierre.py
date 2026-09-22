import os
import unittest

import importador_cierre as ic

_EJEMPLO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ejemplo.json")


class LoaderTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if os.path.exists(_EJEMPLO):
            cls.statements = ic.load_elements(_EJEMPLO)
        else:
            cls.statements = None

    @unittest.skipUnless(os.path.exists(_EJEMPLO), "falta ejemplo.json")
    def test_doble_codificacion(self):
        self.assertIsInstance(self.statements, list)
        self.assertTrue(len(self.statements) >= 1)

    @unittest.skipUnless(os.path.exists(_EJEMPLO), "falta ejemplo.json")
    def test_a1(self):
        self.assertEqual(ic.a1(4, 3), "D5")
        self.assertEqual(ic.a1(25, 26), "AA26")
        self.assertEqual(ic.a1(0, 0), "A1")

    @unittest.skipUnless(os.path.exists(_EJEMPLO), "falta ejemplo.json")
    def test_deteccion_formulas(self):
        recs = ic.extract_formulas(self.statements[0])
        self.assertGreater(len(recs), 100)
        for r in recs[:50]:
            self.assertTrue(r["a1"])
            self.assertTrue(r["name"])
            self.assertTrue(len(r["text"]) > 0)

    @unittest.skipUnless(os.path.exists(_EJEMPLO), "falta ejemplo.json")
    def test_temporalidad(self):
        recs = ic.extract_formulas(self.statements[0])
        with_dia = [r for r in recs if r["temp"] == "DIA"]
        self.assertGreater(len(with_dia), 50)

    @unittest.skipUnless(os.path.exists(_EJEMPLO), "falta ejemplo.json")
    def test_refs_y_relaciones(self):
        recs = ic.extract_formulas(self.statements[0])
        self.assertTrue(all(isinstance(r["refs"], set) for r in recs))
        # la primera fila (DIA) y su SEMANA comparten la misma $CD fuente
        a = recs[0]
        b = next(r for r in recs if r["a1"] == "F6")
        self.assertTrue(a["refs"] & b["refs"] or ic.touches(a, b))

    @unittest.skipUnless(os.path.exists(_EJEMPLO), "falta ejemplo.json")
    def test_integridad(self):
        recs = ic.extract_formulas(self.statements[0])
        i_edit = next(i for i, r in enumerate(recs) if r["a1"] == "E6")
        touches_antes = set(ic.touching_indices(recs, i_edit))
        self.assertGreater(len(touches_antes), 0)
        # simular edición: añade una fuente nueva
        recs[i_edit]["refs"] |= {"$CD_PRUEBA"}
        recs[i_edit]["modificada"] = True
        # todas las que se tocaban pasan a requerir actualización
        for j in touches_antes:
            if j != i_edit:
                recs[j]["estado"] = "REQUIERE ACTUALIZACIÓN"
        pending = [r for r in recs if r["estado"] == "REQUIERE ACTUALIZACIÓN"]
        self.assertGreaterEqual(len(pending), len(touches_antes) - 1)


if __name__ == "__main__":
    unittest.main()