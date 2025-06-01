import sys
import os
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTreeWidget, QTreeWidgetItem, QLabel, QLineEdit, QMessageBox, QAction, QComboBox, QHeaderView, QSpacerItem, QSizePolicy
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.patches import Rectangle

import dicom_utils.reader as dr
import dicom_utils.export_excel as ex
import efs_converter.DCM2EFS as ec

class DicomEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Editor de Tags DICOM (RTPLAN)")
        self.resize(1200, 1080)
        # Menu
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("Arquivo")
        action_save_as = QAction("Salvar Como...", self)
        action_save_as.triggered.connect(self.on_save_as)
        file_menu.addAction(action_save_as)
        action_export_efs = QAction("Exportar EFS", self)
        action_export_efs.triggered.connect(self.export_efs)
        file_menu.addAction(action_export_efs)
        # Layout principal
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        # Painel esquerdo
        left_panel = QVBoxLayout()
        layout.addLayout(left_panel, 60)
        self.btn_open = QPushButton("Abrir RTPLAN DICOM")
        self.btn_open.clicked.connect(self.open_dicom)
        left_panel.addWidget(self.btn_open)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["Tag", "VR", "Name", "Value"])
        self.tree.setColumnWidth(0, 300)
        header = self.tree.header()
        header.setSectionResizeMode(0, QHeaderView.Fixed)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        self.tree.itemClicked.connect(self.on_item_selected)
        left_panel.addWidget(self.tree)
        # Painel direito
        right_panel = QVBoxLayout()
        layout.addLayout(right_panel, 40)
        self.lbl_tag = QLabel("Tag:")
        self.edit_tag = QLineEdit(); self.edit_tag.setReadOnly(True)
        self.lbl_vr = QLabel("VR:")
        self.edit_vr = QLineEdit(); self.edit_vr.setReadOnly(True)
        self.lbl_name = QLabel("Name:")
        self.edit_name = QLineEdit(); self.edit_name.setReadOnly(True)
        self.lbl_value = QLabel("Valor atual:")
        self.edit_value = QLineEdit(); self.edit_value.setReadOnly(True)
        self.btn_save_value = QPushButton("Salvar valor")
        self.btn_save_value.clicked.connect(self.save_value)
        self.btn_save_value.setEnabled(False)
        for w in [self.lbl_tag, self.edit_tag, self.lbl_vr, self.edit_vr, self.lbl_name, self.edit_name, self.lbl_value, self.edit_value, self.btn_save_value]:
            right_panel.addWidget(w)
        # Modelo MLC
        mlc_model_layout = QHBoxLayout()
        mlc_model_layout.addWidget(QLabel("Modelo MLC:"))
        self.mlc_model_combo = QComboBox()
        self.mlc_model_combo.addItem("Agility (5 mm)", 5)
        self.mlc_model_combo.addItem("MLCi2 (10 mm)", 10)
        self.mlc_model_combo.currentIndexChanged.connect(lambda _: self.update_mlc_view())
        mlc_model_layout.addWidget(self.mlc_model_combo)
        right_panel.addLayout(mlc_model_layout)
        # Parâmetros CP
        top_params_layout = QHBoxLayout()
        left_params_layout = QVBoxLayout()
        font_cp_small = QFont(); font_cp_small.setPointSize(12)
        self.lbl_gantry = QLabel("Gantry: N/A"); self.lbl_gantry.setFont(font_cp_small)
        self.lbl_collimator = QLabel("Collimator: N/A"); self.lbl_collimator.setFont(font_cp_small)
        self.lbl_table = QLabel("Table: N/A"); self.lbl_table.setFont(font_cp_small)
        self.lbl_mu = QLabel("MU: N/A"); self.lbl_mu.setFont(font_cp_small)
        for lbl in [self.lbl_gantry, self.lbl_collimator, self.lbl_table, self.lbl_mu]:
            lbl.setAlignment(Qt.AlignLeft)
            left_params_layout.addWidget(lbl)
        top_params_layout.addLayout(left_params_layout)
        spacer = QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        top_params_layout.addItem(spacer)
        self.lbl_fraction = QLabel("Fraction: N/A")
        font_fraction = QFont(); font_fraction.setPointSize(14)
        self.lbl_fraction.setFont(font_fraction); self.lbl_fraction.setAlignment(Qt.AlignRight)
        top_params_layout.addWidget(self.lbl_fraction)
        right_panel.addLayout(top_params_layout)
        right_panel.addWidget(QLabel("— Visualizador de MLC e Jaws —", alignment=Qt.AlignCenter))
        # Navegação CP
        mlc_nav_layout = QHBoxLayout()
        mlc_nav_layout.addWidget(QLabel("Feixe:"))
        self.beam_combo = QComboBox(); self.beam_combo.currentIndexChanged.connect(self.on_beam_changed)
        mlc_nav_layout.addWidget(self.beam_combo)
        self.prev_cp_btn = QPushButton("Anterior CP"); self.prev_cp_btn.clicked.connect(self.on_prev_cp); self.prev_cp_btn.setEnabled(False)
        mlc_nav_layout.addWidget(self.prev_cp_btn)
        self.next_cp_btn = QPushButton("Próximo CP"); self.next_cp_btn.clicked.connect(self.on_next_cp); self.next_cp_btn.setEnabled(False)
        mlc_nav_layout.addWidget(self.next_cp_btn)
        self.cp_label = QLabel("CP: 0/0"); self.cp_label.setFont(font_fraction)
        mlc_nav_layout.addWidget(self.cp_label)
        right_panel.addLayout(mlc_nav_layout)
        # Botões Excel
        btns_excel_layout = QHBoxLayout()
        self.btn_export_excel = QPushButton("Exportar CPs para Excel"); self.btn_export_excel.clicked.connect(self.export_control_points_to_excel)
        self.btn_import_excel = QPushButton("Importar CPs do Excel"); self.btn_import_excel.clicked.connect(self.import_control_points_from_excel)
        btns_excel_layout.addWidget(self.btn_export_excel); btns_excel_layout.addWidget(self.btn_import_excel)
        right_panel.addLayout(btns_excel_layout)
        # Canvas
        self.fig = Figure(figsize=(5, 5)); self.canvas = FigureCanvas(self.fig); self.ax = self.fig.add_subplot(111)
        right_panel.addWidget(self.canvas)
        self.lbl_obs = QLabel("OBS: "); self.lbl_obs.setFont(font_fraction); self.lbl_obs.setAlignment(Qt.AlignLeft)
        right_panel.addWidget(self.lbl_obs)
        # Estados
        self.dataset = None; self.dicom_path = None; self.current_item = None; self.current_element = None
        self.current_beam_idx = 0; self.current_cp_idx = 0; self._excel_path = None

    def open_dicom(self):
        path, _ = QFileDialog.getOpenFileName(self, "Selecione RTPLAN DICOM", "", "DICOM (*.dcm *.DCM)")
        if not path: return
        try:
            self.dataset = dr.open_dicom_file(path)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao ler DICOM:\n{e}"); return
        self.dicom_path = path; self.tree.clear()
        dr.populate_tree(self.dataset, self.tree.invisibleRootItem()); self.setWindowTitle(f"Editor DICOM — {path}")
        self.init_beam_cp_view()

    def on_item_selected(self, item, column):
        self.current_item = item; elem = getattr(item, "data_element", None)
        if elem is None or elem.VR == "SQ":
            self.current_element = None; self.edit_tag.setText(""); self.edit_vr.setText(""); self.edit_name.setText(item.text(2)); self.edit_value.setText(""); self.edit_value.setReadOnly(True); self.btn_save_value.setEnabled(False)
            return
        self.current_element = elem; self.edit_tag.setText(f"({elem.tag.group:04X},{elem.tag.element:04X})"); self.edit_vr.setText(elem.VR); self.edit_name.setText(elem.name)
        val = elem.value; val_str = "<bytes...>" if isinstance(val, bytes) and len(val) > 64 else str(val)
        self.edit_value.setText(val_str); self.edit_value.setReadOnly(False); self.btn_save_value.setEnabled(True)

    def save_value(self):
        if self.current_element is None: return
        new_str = self.edit_value.text().strip()
        try:
            dr.save_data_element(self.current_element, new_str)
        except Exception as e:
            QMessageBox.warning(self, "Erro", f"Não possível converter:\n{e}"); return
        self.current_item.setText(3, new_str)
        QMessageBox.information(self, "OK", "Valor atualizado em memória.\nUse “Salvar Como...” para gravar.")

    def on_save_as(self):
        if self.dataset is None: QMessageBox.warning(self, "Atenção", "Nenhum arquivo aberto."); return
        path_out, _ = QFileDialog.getSaveFileName(self, "Salvar DICOM", "", "DICOM (*.dcm)")
        if not path_out: return
        try:
            dr.save_dicom_file(self.dataset, path_out)
            QMessageBox.information(self, "Arquivo salvo", f"Gravado em:\n{path_out}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Não foi possível salvar:\n{e}")

    def init_beam_cp_view(self):
        self.beam_combo.clear(); self.current_beam_idx = 0; self.current_cp_idx = 0
        beams = dr.get_beams(self.dataset)
        if not beams:
            QMessageBox.warning(self, "Atenção", "Sem feixes neste arquivo."); self.beam_combo.setEnabled(False); self.prev_cp_btn.setEnabled(False); self.next_cp_btn.setEnabled(False); self.cp_label.setText("CP: 0/0"); self.ax.clear(); self.ax.set_title("Sem dados"); self.canvas.draw(); return
        for idx, beam in enumerate(beams):
            beam_num = getattr(beam, "BeamNumber", idx + 1); beam_name = getattr(beam, "BeamName", "")
            self.beam_combo.addItem(f"{beam_num} {beam_name}", idx)
        self.beam_combo.setEnabled(True); self.beam_combo.setCurrentIndex(0); self.prev_cp_btn.setEnabled(True); self.next_cp_btn.setEnabled(True)
        self.update_mlc_view()

    def on_beam_changed(self, index):
        self.current_beam_idx = index; self.current_cp_idx = 0; self.update_mlc_view()

    def on_prev_cp(self):
        beams = dr.get_beams(self.dataset); cps = dr.get_control_points(beams[self.current_beam_idx])
        if self.current_cp_idx > 0: self.current_cp_idx -= 1; self.update_mlc_view()

    def on_next_cp(self):
        beams = dr.get_beams(self.dataset); cps = dr.get_control_points(beams[self.current_beam_idx])
        if self.current_cp_idx < len(cps)-1: self.current_cp_idx += 1; self.update_mlc_view()

    def update_mlc_view(self):
        self.ax.clear(); beams = dr.get_beams(self.dataset)
        if not beams: self.ax.set_title("Sem dados"); self.canvas.draw(); return
        beam = beams[self.current_beam_idx]; cps = dr.get_control_points(beam); total_cps = len(cps)
        self.cp_label.setText(f"CP: {self.current_cp_idx+1}/{total_cps}")
        cp = cps[self.current_cp_idx]; bl_seq = dr.get_bl_seq(cp)
        if not bl_seq: self.ax.set_title("Sem BeamLimitingDevicePositionSequence"); self.canvas.draw(); self.lbl_obs.setText("OBS: Sem BL seq"); return
        gantry = getattr(cp, "GantryAngle", "N/A"); collim = getattr(cp, "BeamLimitingDeviceAngle", "N/A")
        table = getattr(cp, "PatientSupportAngle", "N/A"); fraction = getattr(cp, "CumulativeMetersetWeight", "N/A")
        mu = getattr(beam, "BeamMeterset", "N/A")
        self.lbl_gantry.setText(f"Gantry: {gantry}°"); self.lbl_collimator.setText(f"Collimator: {collim}°"); self.lbl_table.setText(f"Table: {table}°"); self.lbl_mu.setText(f"MU: {mu}"); self.lbl_fraction.setText(f"Fraction: {fraction}")
        mlc_item = dr.find_mlc_item(bl_seq)
        if not mlc_item: self.ax.set_title("Nenhum MLC"); self.canvas.draw(); self.lbl_obs.setText("OBS: Sem MLC"); return
        leaf_positions = dr.get_mlc_positions(mlc_item); N2 = len(leaf_positions); N = N2//2
        left_positions = leaf_positions[:N]; right_positions = leaf_positions[N:]
        x_jaws = dr.find_jaw_positions(bl_seq, axis='X'); y_jaws = dr.find_jaw_positions(bl_seq, axis='Y')
        obs_msgs = []; 
        if not x_jaws: obs_msgs.append("Sem X")
        if not y_jaws: obs_msgs.append("Sem Y")
        self.lbl_obs.setText("OBS: " + "; ".join(obs_msgs) if obs_msgs else "OBS: Colimadores OK")
        xmin, xmax = -200.0, 200.0
        thickness = float(self.mlc_model_combo.currentData()); total_height = N * thickness; half_height = total_height/2.0
        for i in range(N):
            left_val = left_positions[i]; right_val = right_positions[i]; y_bottom = i*thickness - half_height
            if left_val > xmin:
                width_left = left_val - xmin
                rect_left = Rectangle((xmin, y_bottom), width_left, thickness, facecolor="lightgray", edgecolor="black"); self.ax.add_patch(rect_left)
            if xmax > right_val:
                width_right = xmax - right_val
                rect_right = Rectangle((right_val, y_bottom), width_right, thickness, facecolor="lightgray", edgecolor="black"); self.ax.add_patch(rect_right)
        if x_jaws:
            xl, xr = x_jaws; xl = max(min(xl,200.0),-200.0); xr = max(min(xr,200.0),-200.0)
            self.ax.axvline(x=xl, color="red", linestyle="--", linewidth=1.5, label="X Jaw L"); self.ax.axvline(x=xr, color="red", linestyle="--", linewidth=1.5, label="X Jaw R")
        if y_jaws:
            y1, y2 = y_jaws
            self.ax.axhline(y=y1, color="blue", linestyle="--", linewidth=1.5, label=f"Y Jaw {y1}"); self.ax.axhline(y=y2, color="blue", linestyle="--", linewidth=1.5, label=f"Y Jaw {y2}")
            new_top = max(half_height, y2); new_bot = min(-half_height, y1); self.ax.set_ylim(new_top, new_bot)
        else:
            self.ax.set_ylim(half_height, -half_height)
        self.ax.set_xlim(xmin, xmax); self.ax.set_xlabel("X (mm)"); self.ax.set_ylabel("Y (mm)")
        beam_label = getattr(beam,"BeamNumber", self.current_beam_idx+1)
        self.ax.set_title(f"Beam {beam_label} · CP {self.current_cp_idx+1}/{total_cps}")
        if x_jaws or y_jaws: self.ax.legend(loc="upper right", fontsize="small")
        self.ax.grid(True, linestyle=":", linewidth=0.5, alpha=0.6); self.canvas.draw()

    def export_control_points_to_excel(self):
        ex.export_to_excel(self.dataset)

    def import_control_points_from_excel(self):
        ex.import_from_excel(self.dataset, self.current_beam_idx)

    def export_efs(self):
        if not self.dataset or not self.dicom_path:
            QMessageBox.warning(self, "Atenção", "Abra um RTPLAN antes de exportar EFS."); return
        output_folder = QFileDialog.getExistingDirectory(self, "Selecione pasta destino EFS")
        if not output_folder: return
        try:
            ec.convert_dcm2efs(self.dicom_path, output_folder)
        except Exception as e:
            QMessageBox.critical(self, "Erro EFS", f"Falha ao gerar EFS:\n{e}"); return
        QMessageBox.information(self, "Exportado", f"EFS gerados em:\n{output_folder}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    editor = DicomEditor()
    editor.show()
    sys.exit(app.exec_())
