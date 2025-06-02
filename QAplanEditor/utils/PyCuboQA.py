import numpy as np
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pydicom
from pydicom.dataset import FileDataset
import datetime
import os

def gerar_volume_com_cubo_mm(mm_x, mm_y, mm_z, pixel_size, slice_thickness):
    """
    Cria um volume cúbico 3D onde há um cubo de água (0 HU)
    centralizado num fundo de ar (–1000 HU).
    O lado do volume é igual ao maior dos mm_x, mm_y, mm_z + 100 mm.
    """
    maior_dim = max(mm_x, mm_y, mm_z)
    lado_volume_mm = maior_dim + 100.0  # adiciona 10 cm de margem

    # Quantos voxels em cada direção do volume
    voxels_z = int(round(lado_volume_mm / slice_thickness))
    voxels_y = int(round(lado_volume_mm / pixel_size))
    voxels_x = int(round(lado_volume_mm / pixel_size))

    # Quantos voxels no cubo de água
    vox_cubo_z = int(round(mm_z / slice_thickness))
    vox_cubo_y = int(round(mm_y / pixel_size))
    vox_cubo_x = int(round(mm_x / pixel_size))

    # Cria volume cúbico preenchido com ar (–1000 HU)
    volume = np.full((voxels_z, voxels_y, voxels_x), -1000, dtype=np.int16)

    # Índices de início para centralizar o cubo
    start_z = (voxels_z - vox_cubo_z) // 2
    start_y = (voxels_y - vox_cubo_y) // 2
    start_x = (voxels_x - vox_cubo_x) // 2

    end_z = start_z + vox_cubo_z
    end_y = start_y + vox_cubo_y
    end_x = start_x + vox_cubo_x

    # Insere cubo de água (0 HU)
    volume[start_z:end_z, start_y:end_y, start_x:end_x] = 0

    voxel_spacing = (slice_thickness, pixel_size, pixel_size)  # Z, Y, X
    return volume, voxel_spacing

def exportar_dicom(volume, voxel_spacing, patient_name, patient_id, output_dir):
    """
    Exporta cada slice axial do volume como um arquivo DICOM no diretório dado,
    usando os tags PatientName, PatientID, StudyInstanceUID, SeriesInstanceUID
    e FrameOfReferenceUID gerados automaticamente.
    Retorna o caminho da pasta de saída.
    """
    dt = datetime.datetime.now()
    study_uid = pydicom.uid.generate_uid()
    series_uid = pydicom.uid.generate_uid()
    frame_uid = pydicom.uid.generate_uid()  # FrameOfReferenceUID

    depth, height, width = volume.shape  # Z, Y, X

    # Garante que o diretório de saída exista
    os.makedirs(output_dir, exist_ok=True)

    for z in range(depth):
        file_meta = pydicom.Dataset()
        file_meta.MediaStorageSOPClassUID = pydicom.uid.CTImageStorage
        file_meta.MediaStorageSOPInstanceUID = pydicom.uid.generate_uid()
        file_meta.ImplementationClassUID = pydicom.uid.generate_uid()

        ds = FileDataset('', {}, file_meta=file_meta, preamble=b"\0" * 128)
        ds.ContentDate = dt.strftime('%Y%m%d')
        ds.ContentTime = dt.strftime('%H%M%S')
        ds.PatientName = patient_name
        ds.PatientID = patient_id
        ds.StudyInstanceUID = study_uid
        ds.SeriesInstanceUID = series_uid
        ds.FrameOfReferenceUID = frame_uid
        ds.SOPInstanceUID = file_meta.MediaStorageSOPInstanceUID
        ds.SOPClassUID = file_meta.MediaStorageSOPClassUID
        ds.Modality = 'CT'
        ds.Rows, ds.Columns = height, width
        ds.InstanceNumber = z + 1
        ds.PixelSpacing = [str(voxel_spacing[1]), str(voxel_spacing[2])]
        ds.SliceThickness = str(voxel_spacing[0])
        ds.ImagePositionPatient = [0.0, 0.0, float(z) * voxel_spacing[0]]
        ds.ImageOrientationPatient = [1, 0, 0, 0, 1, 0]
        ds.SamplesPerPixel = 1
        ds.PhotometricInterpretation = "MONOCHROME2"
        ds.PixelRepresentation = 1
        ds.HighBit = 15
        ds.BitsStored = 16
        ds.BitsAllocated = 16
        ds.RescaleIntercept = 0
        ds.RescaleSlope = 1
        ds.WindowCenter = 0
        ds.WindowWidth = 4000
        ds.PixelData = volume[z, :, :].tobytes()

        filename = os.path.join(output_dir, f'slice_{z:03}.dcm')
        ds.save_as(filename)

    return output_dir

def atualizar_imagem():
    """
    Lê valores da GUI, gera volume cúbico e exibe cortes axial, coronal e sagital.
    """
    try:
        mm_x = float(entry_lx.get())
        mm_y = float(entry_ly.get())
        mm_z = float(entry_lz.get())
        pixel_size = float(entry_ps.get())
        thickness = float(entry_st.get())

        if mm_x <= 0 or mm_y <= 0 or mm_z <= 0 or pixel_size <= 0 or thickness <= 0:
            raise ValueError

        volume, spacing = gerar_volume_com_cubo_mm(mm_x, mm_y, mm_z, pixel_size, thickness)

        depth, height, width = volume.shape
        z_mid = depth // 2
        y_mid = height // 2
        x_mid = width // 2

        slice_axial   = volume[z_mid, :, :]   # plano XY
        slice_coronal = volume[:, y_mid, :]   # plano XZ
        slice_sagital = volume[:, :, x_mid]   # plano YZ

        axs[0].clear()
        axs[0].imshow(slice_axial, cmap='gray', vmin=-1000, vmax=500)
        axs[0].set_title(f'Axial (Z={z_mid})\nDim Cubo: {mm_x:.0f}×{mm_y:.0f}×{mm_z:.0f} mm')
        axs[0].axis('off')

        axs[1].clear()
        axs[1].imshow(slice_coronal, cmap='gray', vmin=-1000, vmax=500)
        axs[1].set_title(f'Coronal (Y={y_mid})\nPixel: {pixel_size} mm, Corte: {thickness} mm')
        axs[1].axis('off')

        axs[2].clear()
        axs[2].imshow(slice_sagital, cmap='gray', vmin=-1000, vmax=500)
        axs[2].set_title(f'Sagital (X={x_mid})')
        axs[2].axis('off')

        canvas.draw()

        global current_volume, current_spacing
        current_volume = volume
        current_spacing = spacing

    except ValueError:
        messagebox.showerror("Erro", "Insira valores válidos e positivos.")

def salvar_dicom():
    """
    Lê PatientName, PatientID e diretório de saída da GUI e chama exportar_dicom,
    gerando também StudyInstanceUID, SeriesInstanceUID e FrameOfReferenceUID automaticamente.
    """
    if current_volume is not None and current_spacing is not None:
        patient_name = entry_patient_name.get().strip()
        patient_id = entry_patient_id.get().strip()
        if not patient_name:
            messagebox.showerror("Erro", "Insira o nome do paciente.")
            return
        if not patient_id:
            messagebox.showerror("Erro", "Insira o Patient ID.")
            return

        # Solicita ao usuário o diretório para salvar
        output_dir = filedialog.askdirectory(title="Selecione o diretório para salvar DICOMs")
        if not output_dir:
            return  # Usuário cancelou

        pasta = exportar_dicom(current_volume, current_spacing, patient_name, patient_id, output_dir)
        messagebox.showinfo("DICOM Exportado", f"Arquivos DICOM salvos em:\n{pasta}")
    else:
        messagebox.showerror("Erro", "Gere primeiro o volume.")

# --- Montagem da GUI ---
root = tk.Tk()
root.title("Simulador de Cubo de Água em TC - Volume Cúbico")

frame_inputs = tk.Frame(root)
frame_inputs.pack(pady=10)

# Dimensões do cubo
tk.Label(frame_inputs, text="Largura X (mm):").grid(row=0, column=0)
entry_lx = tk.Entry(frame_inputs, width=6)
entry_lx.insert(0, "300")
entry_lx.grid(row=0, column=1)

tk.Label(frame_inputs, text="Altura Y (mm):").grid(row=0, column=2)
entry_ly = tk.Entry(frame_inputs, width=6)
entry_ly.insert(0, "300")
entry_ly.grid(row=0, column=3)

tk.Label(frame_inputs, text="Comprimento Z (mm):").grid(row=0, column=4)
entry_lz = tk.Entry(frame_inputs, width=6)
entry_lz.insert(0, "300")
entry_lz.grid(row=0, column=5)

# Resolução
tk.Label(frame_inputs, text="Tamanho do pixel (mm):").grid(row=1, column=0)
entry_ps = tk.Entry(frame_inputs, width=6)
entry_ps.insert(0, "1.0")
entry_ps.grid(row=1, column=1)

tk.Label(frame_inputs, text="Espessura do corte (mm):").grid(row=1, column=2)
entry_st = tk.Entry(frame_inputs, width=6)
entry_st.insert(0, "3.0")
entry_st.grid(row=1, column=3)

# Campos de paciente
tk.Label(frame_inputs, text="Patient Name:").grid(row=2, column=0)
entry_patient_name = tk.Entry(frame_inputs, width=15)
entry_patient_name.insert(0, "Phantom^Cube")
entry_patient_name.grid(row=2, column=1)

tk.Label(frame_inputs, text="Patient ID:").grid(row=2, column=2)
entry_patient_id = tk.Entry(frame_inputs, width=10)
entry_patient_id.insert(0, "123456")
entry_patient_id.grid(row=2, column=3)

# Botões
btn_gerar = tk.Button(root, text="Gerar e Visualizar", command=atualizar_imagem)
btn_gerar.pack(pady=5)

btn_exportar = tk.Button(root, text="Exportar como DICOM", command=salvar_dicom)
btn_exportar.pack(pady=5)

# Área de plotagem
fig, axs = plt.subplots(1, 3, figsize=(12, 4))
canvas = FigureCanvasTkAgg(fig, master=root)
canvas.get_tk_widget().pack()

current_volume = None
current_spacing = None

# Gera a primeira visualização automaticamente
atualizar_imagem()
root.mainloop()
