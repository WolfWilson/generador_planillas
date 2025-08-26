#!/usr/bin/env python
# coding: utf-8
"""
GUI – Generador de Planillas de Horarios (PySide6)
- Carga robusta de QSS (QIODevice + soporte PyInstaller _MEIPASS)
- Sugerencia de ruta de guardado en Documentos
- Validación de horas HH:MM,HH:MM
"""

from __future__ import annotations

import os
import sys
from datetime import datetime
from typing import Tuple

from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QSpinBox,
    QGroupBox,
    QComboBox,
)
from PySide6.QtCore import (
    Qt,
    QFile,
    QTextStream,
    QIODevice,
    QStandardPaths,
)

# Importa tu lógica (debes tener estos módulos en el mismo proyecto)
from generar_planilla_word import generar_planilla_word, parse_date_list, parse_notes


class PlanillaApp(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Generador de Planillas de Horarios")
        self.init_ui()
        self.load_style()
        self.resize(640, 540)

    def init_ui(self) -> None:
        layout = QVBoxLayout(self)

        # --- Datos Principales ---
        group_main = QGroupBox("Datos Principales")
        main_layout = QVBoxLayout()

        self.nombre_edit = QLineEdit("Benitez Wilson")
        main_layout.addWidget(QLabel("Apellido y Nombre:"))
        main_layout.addWidget(self.nombre_edit)

        self.oficina_edit = QLineEdit("CPI")
        main_layout.addWidget(QLabel("Oficina / Sector:"))
        main_layout.addWidget(self.oficina_edit)

        self.empleado_edit = QLineEdit("32.746.256")
        main_layout.addWidget(QLabel("Legajo / DNI:"))
        main_layout.addWidget(self.empleado_edit)

        # Mes y Año
        now = datetime.now()
        mes_anio_layout = QHBoxLayout()

        self.mes_combo = QComboBox()
        nombres_mes = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
        ]
        self.mes_combo.addItems(nombres_mes)
        self.mes_combo.setCurrentIndex(now.month - 1)
        mes_anio_layout.addWidget(QLabel("Mes:"))
        mes_anio_layout.addWidget(self.mes_combo)

        self.anio_spin = QSpinBox()
        self.anio_spin.setRange(2020, 2100)
        self.anio_spin.setValue(now.year)
        mes_anio_layout.addWidget(QLabel("Año:"))
        mes_anio_layout.addWidget(self.anio_spin)

        main_layout.addLayout(mes_anio_layout)
        group_main.setLayout(main_layout)
        layout.addWidget(group_main)

        # --- Opciones Avanzadas ---
        group_options = QGroupBox("Opciones (opcional)")
        options_layout = QVBoxLayout()

        self.hora_m_edit = QLineEdit("06:30,13:00")
        options_layout.addWidget(QLabel("Horario Mañana (entrada,salida):"))
        options_layout.addWidget(self.hora_m_edit)

        self.hora_t_edit = QLineEdit("16:00,19:00")
        options_layout.addWidget(QLabel("Horario Tarde (entrada,salida):"))
        options_layout.addWidget(self.hora_t_edit)

        self.extras_dow_edit = QLineEdit("1,3")
        options_layout.addWidget(QLabel("Días con extras (0=Lun, 1=Mar, ... 6=Dom):"))
        options_layout.addWidget(self.extras_dow_edit)

        self.feriados_edit = QLineEdit()
        self.feriados_edit.setPlaceholderText("Ej: 2025-09-11,2025-09-15")
        options_layout.addWidget(QLabel("Feriados (YYYY-MM-DD, ...):"))
        options_layout.addWidget(self.feriados_edit)

        self.notas_edit = QLineEdit()
        self.notas_edit.setPlaceholderText("Ej: 16:LICENCIA,17:CAPACITACIÓN")
        options_layout.addWidget(QLabel("Notas por día (dia:texto, ...):"))
        options_layout.addWidget(self.notas_edit)

        group_options.setLayout(options_layout)
        layout.addWidget(group_options)

        # --- Archivo de Salida ---
        group_output = QGroupBox("Archivo de Salida")
        output_layout = QHBoxLayout()
        self.out_path_edit = QLineEdit()
        self.out_path_edit.setPlaceholderText("Seleccione la ruta de salida...")
        output_layout.addWidget(self.out_path_edit)

        browse_btn = QPushButton("...")
        browse_btn.clicked.connect(self.browse_output_file)
        output_layout.addWidget(browse_btn)
        group_output.setLayout(output_layout)
        layout.addWidget(group_output)

        # --- Botón de Generar ---
        self.generate_btn = QPushButton("Generar Planilla")
        self.generate_btn.clicked.connect(self.generate)
        layout.addWidget(self.generate_btn)

    def center(self) -> None:
        """Centra la ventana en la pantalla."""
        qr = self.frameGeometry()
        cp = self.screen().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def _parse_horas(self, text: str) -> Tuple[str, str]:
        """Valida formato HH:MM,HH:MM y retorna (entrada, salida)."""
        parts = [p.strip() for p in text.split(",")]
        if len(parts) != 2 or any(":" not in p for p in parts):
            raise ValueError("Formato de hora inválido. Use HH:MM,HH:MM (ej: 06:30,13:00)")
        h1, h2 = parts[0], parts[1]
        # Validación simple de hh:mm
        for hhmm in (h1, h2):
            hhmm_parts = hhmm.split(":")
            if len(hhmm_parts) != 2:
                raise ValueError(f"Hora inválida: {hhmm}")
            hh, mm = hhmm_parts[0], hhmm_parts[1]
            if not (hh.isdigit() and mm.isdigit()):
                raise ValueError(f"Hora inválida: {hhmm}")
            if not (0 <= int(hh) <= 23 and 0 <= int(mm) <= 59):
                raise ValueError(f"Hora fuera de rango: {hhmm}")
        return h1, h2

    def load_style(self) -> None:
        """Carga la hoja de estilos QSS, compatible con ejecución normal y PyInstaller."""
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            # Soporte PyInstaller (_MEIPASS)
            if hasattr(sys, "_MEIPASS"):
                base_dir = sys._MEIPASS  # type: ignore[attr-defined]

            style_path = os.path.join(base_dir, "styles", "style.qss")
            style_file = QFile(style_path)
            flags = QIODevice.OpenModeFlag.ReadOnly | QIODevice.OpenModeFlag.Text
            if not style_file.open(flags):
                print(f"⚠ No se pudo abrir el archivo de estilos: {style_path}")
                return

            stream = QTextStream(style_file)
            qss = stream.readAll()
            style_file.close()

            self.setStyleSheet(qss)
        except Exception as exc:  # noqa: BLE001
            print(f"⚠ Error cargando estilos: {exc}")

    def browse_output_file(self) -> None:
        """Propone guardar en Documentos con nombre sugerido."""
        docs = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        mes_nombre = self.mes_combo.currentText().upper()
        anio = self.anio_spin.value()
        nombre_sugerido = os.path.join(docs, f"PLANILLA-{mes_nombre}-{anio}.docx")

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Guardar Planilla",
            nombre_sugerido,
            "Documentos Word (*.docx)",
        )
        if path:
            self.out_path_edit.setText(path)

    def generate(self) -> None:
        # Recopilar datos de la UI
        nombre = self.nombre_edit.text().strip()
        oficina = self.oficina_edit.text().strip()
        empleado = self.empleado_edit.text().strip()
        out_path = self.out_path_edit.text().strip()

        if not all([nombre, oficina, empleado, out_path]):
            QMessageBox.warning(
                self,
                "Campos incompletos",
                "Por favor, complete todos los datos principales y seleccione una ruta de salida.",
            )
            return

        try:
            mes = self.mes_combo.currentIndex() + 1
            anio = self.anio_spin.value()

            h_m_entrada, h_m_salida = self._parse_horas(self.hora_m_edit.text())
            h_t_entrada, h_t_salida = self._parse_horas(self.hora_t_edit.text())

            extras_str = self.extras_dow_edit.text().strip()
            extras = tuple(int(x) for x in extras_str.split(",") if x.strip()) if extras_str else ()

            feriados = parse_date_list(self.feriados_edit.text())
            notas = parse_notes(self.notas_edit.text())

            # Llamada principal a tu generador Word
            generar_planilla_word(
                out_path=out_path,
                nombre=nombre,
                oficina=oficina,
                empleado=empleado,
                mes=mes,
                anio=anio,
                hora_maniana=(h_m_entrada, h_m_salida),
                hora_tarde=(h_t_entrada, h_t_salida),
                extras_dow=extras,
                notas_por_dia=notas,
                feriados=feriados,
            )

            QMessageBox.information(self, "Éxito", f"Planilla generada correctamente en:\n{out_path}")

        except Exception as e:  # noqa: BLE001
            QMessageBox.critical(self, "Error", f"Ocurrió un error al generar la planilla:\n{e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PlanillaApp()
    window.show()
    window.center()
    sys.exit(app.exec())
