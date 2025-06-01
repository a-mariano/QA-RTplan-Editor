# My DICOM RTPLAN Editor

Este projeto fornece uma interface gráfica para:
- Visualizar e editar tags DICOM de arquivos RTPLAN.
- Visualizar posições de MLC e Jaws para cada Control Point.
- Exportar e importar parâmetros de Control Points para/desde Excel.
- Exportar cada feixe para arquivos `.efs`.

## Estrutura de Arquivos

```
my-dicom-editor/
│
├── README.md
├── requirements.txt
├── main_window.py         # Janela principal da GUI
│
├── dicom_utils/
│   ├── __init__.py
│   ├── reader.py          # Funções para abrir e navegar no DICOM
│   └── export_excel.py    # Exportação/importação de CPs para Excel
│
├── efs_converter/
│   ├── __init__.py
│   └── converter.py       # Conversão de RTPLAN DICOM para arquivos .efs
│
└── utils/
    ├── __init__.py
    └── helpers.py         # Funções auxiliares gerais
```

## Instalação

1. Clone o repositório:
   ```
   git clone https://github.com/seu-usuario/my-dicom-editor.git
   cd my-dicom-editor
   ```

2. Crie um ambiente virtual (opcional, mas recomendado):
   ```
   python -m venv venv
   source venv/bin/activate   # Linux/Mac
   venv\Scripts\activate    # Windows
   ```

3. Instale as dependências:
   ```
   pip install -r requirements.txt
   ```

## Uso

Execute o arquivo principal:

```
python main_window.py
```

- Clique em **"Abrir RTPLAN DICOM"** para selecionar um arquivo `.dcm`.
- A árvore exibirá todas as tags. Selecione uma tag para editar.
- Use **"Salvar Como..."** no menu Arquivo para gravar alterações em disco.
- Use o botão **"Exportar CPs para Excel"** para gerar uma planilha com os parâmetros de cada Control Point.
- Após editar e fechar o Excel, clique em **"Importar CPs do Excel"** para atualizar o DICOM em memória.
- Use **"Exportar EFS"** no menu Arquivo para gerar arquivos `.efs` por feixe.

## Licença

Este projeto está licenciado sob a MIT License.
