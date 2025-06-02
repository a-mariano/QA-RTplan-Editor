# QA-RTplan-Editor

Este repositório contém o **QA-RTplan-Editor**, uma aplicação em Python com interface gráfica (PyQt5) para:

- Abrir, visualizar e editar **RTPLAN DICOM** de um plano de radioterapia.
- Navegar pelos beams e control points, com visualização de **MLC** e **jaw positions**.
- Exportar e importar parâmetros de control points via **Excel**.
- Exportar para o formato **.efs** (usando o conversor DCM3EFS.py sem modificações).
- Gerar um **CT Phantom** (cubo de água em fundo de ar) por meio da interface original do `PyCuboQA.py` (Tkinter).
- Atualizar, em memória, o RTPLAN para referenciar esse novo CT (ajustando PatientName, PatientID, Study/Series UIDs, FrameOfReferenceUID e sequências de referência).
- Salvar as modificações do RTPLAN em disco quando desejado.

---

## Índice

1. [Visão Geral](#visão-geral)  
2. [Funcionalidades Principais](#funcionalidades-principais)  
   1. [Edição de Tags DICOM (RTPLAN)](#edição-de-tags-dicom-rtplan)  
   2. [Visualização de MLC e Jaws](#visualização-de-mlc-e-jaws)  
   3. [Navegação entre Beams e Control Points](#navegação-entre-beams-e-control-points)  
   4. [Exportar/Importar Control Points em Excel](#exportarimportar-control-points-em-excel)  
   5. [Exportar para EFS](#exportar-para-efs)  
   6. [Gerar CT Phantom (PyCuboQA)](#gerar-ct-phantom-pycuboqa)  
   7. [Atualizar RTPLAN para Nova CT em Memória](#atualizar-rtplan-para-nova-ct-em-memória)  
3. [Estrutura do Projeto](#estrutura-do-projeto)  
4. [Instalação e Dependências](#instalação-e-dependências)  
5. [Como Executar](#como-executar)  
6. [Menu e Uso da GUI](#menu-e-uso-da-gui)  
7. [Licença](#licença)  

---

## Visão Geral

O **QA-RTplan-Editor** foi desenvolvido para facilitar a análise, verificação e edição de arquivos RTPLAN DICOM utilizados em radioterapia. Com ele, é possível:

- Inspecionar e modificar **qualquer Data Element** (exceto sequências) diretamente na interface.
- Visualizar, slice a slice, o comportamento de **MLC** (Multileaf Collimator) e **jaws** em cada Control Point, com cores e limites fixos.
- Exportar e importar tabelas de Control Points em Excel, formatadas em colunas (cada coluna = um Control Point), incluindo posições de todas as lâminas MLC (esquerda e direita separadas) e novas colunas caso sejam adicionados CPs.
- Gerar e abrir um **CT Phantom** (cubo de água centralizado em fundo de ar) usando a interface Tkinter original (`PyCuboQA.py`), sem replicar diálogos—ao clicar no menu, a janela do PyCuboQA surge com visualização de cortes e parâmetros.
- Adaptar o RTPLAN atual para referenciar esse CT (em memória), copiando `PatientName`, `PatientID`, `StudyInstanceUID`, `SeriesInstanceUID` e `FrameOfReferenceUID` diretamente do primeiro slice do CT; atualizar também `ReferencedStudySequence`, `ReferencedSeriesSequence` e `ReferencedFrameOfReferenceSequence`.
- Exportar o RTPLAN modificado para disco (por “Salvar Como…”) ou gerar arquivos `.efs` — o conversor `DCM3EFS.py` também está incluído sem alterações, e pode ser chamado diretamente pelo menu “Arquivo → Exportar EFS”.

O resultado final é uma ferramenta única para QA de planos de radioterapia, edição de tags e integração/fusão com imagens de referência.

---

## Funcionalidades Principais

### Edição de Tags DICOM (RTPLAN)

- **Árvore hierárquica** de todos os Data Elements (incluindo sequences).  
- Clicar em qualquer tag exibe **Group, Element, VR, Nome** e valor atual.  
- Para elementos de VR diferente de SQ, é possível **editar o valor** diretamente e salvá-lo em memória.  
- Comando “Salvar Como...” grava o DICOM modificado em disco.

### Visualização de MLC e Jaws

- Para cada **Control Point**, desenha:
  - Retângulos horizontais indicando a porção “aberta” do MLC (ambos bancos de lâminas, com offset centralizado, espessura definida no combo de modelo).
  - **Jaws X**: duas linhas verticais (vermelhas) nos valores `Leaf/Jaw Positions` axis X.  
  - **Jaws Y**: duas linhas horizontais (azuis) nos valores `Leaf/Jaw Positions` axis Y.  
- O eixo X é fixo de **−200 mm a +200 mm** (tamanho máximo de campo).  
- Constrói dinamicamente a altura total (número de lâminas × espessura do modelo: “Agility 5 mm” ou “MLCi2 10 mm”).  
- Em “OBS” (abaixo do canvas), indica se não há jaw X e/ou jaw Y naquele CP.  
- Religa a visualização a cada mudança de **beam** ou **control point**.

### Navegação entre Beams e Control Points

- Combo “Feixe” lista todos os elementos de `RTBeamSequence` (exibindo `BeamNumber` e `BeamName`).  
- Botões “Anterior CP” / “Próximo CP” percorrem todos os CPs do beam selecionado.  
- Rótulo “CP: X/Y” mostra índice atual e total.  
- Ao trocar de beam ou CP, a árvore de DICOM permanece, mas o visualizador de MLC/Jaws é atualizado automaticamente.

### Exportar/Importar Control Points em Excel

- **Exportar CPs para Excel**: gera um arquivo `*.xlsx` em que cada **coluna** corresponde a um control point.  
  - A primeira coluna (largura dobrada) lista **nomes das variáveis** (ex.: `GantryAngle`, `BeamMeterset`, `Leaf_Jaw_Positions[0] Left`, etc.).  
  - Para cada CP, há colunas contendo:
    - `GantryAngle`  
    - `BeamLimitingDeviceAngle` (colimador)  
    - `PatientSupportAngle` (mesa)  
    - `CumulativeMetersetWeight` (fração de dose)  
    - `BeamMeterset` (MU)  
    - Todas as posições de cada lâmina MLC (cada lâmina em célula separada, numeradas a partir de 1; indicando também “Left” ou “Right”).  
  - Se o plano tiver **mais de um campo (beam)**, cada beam é esportado em **planilhas diferentes** do mesmo arquivo.  
- **Importar CPs do Excel**: ao fechar o Excel, o aplicativo lê o arquivo e atualiza todos os CPs do beam selecionado (ou adiciona novos CPs, copiando parâmetros do último CP e substituindo apenas os valores fornecidos no Excel).  
- Essa troca bidirecional facilita a edição em lote de dezenas/hundreds de CPs e lamelas.

### Exportar para EFS

- Comando “Arquivo → Exportar EFS”: chama o script original `DCM3EFS.py` (sem nenhuma modificação) para converter o DICOM RTPLAN em arquivos `.efs`.  
- O usuário escolhe a pasta de destino; o conversor gera, para cada CP/beam, o `.efs` correspondente.  
- Caso falhe algum CP, exibe **mensagem de erro** detalhando o motivo.

### Gerar CT Phantom (PyCuboQA)

- Menu “CT → Gerar Novo CT Phantom” abre **exclusivamente** a interface Tkinter original de `PyCuboQA.py` (terminal separado), que contém:
  1. Campos para inserir **dimensões (X, Y, Z em mm)**, **pixel size**, **slice thickness**, **PatientName** e **PatientID**.  
  2. Visualizador de cortes (usando Matplotlib no Tkinter), mostrando o cubo de água centralizado em fundo de ar.  
  3. Botão para selecionar diretório e exportar cada slice axial como DICOM.  
- Esse código foi mantido **sem alterações** (importado via `subprocess.Popen`), garantindo que a interface original seja preservada.

### Atualizar RTPLAN para Nova CT em Memória

- Menu “CT → Atualizar RTPLAN com CT” (sem pedir “Salvar Como…”):
  1. Abre diálogo para selecionar **pasta** contendo os DICOMs do CT recém-gerado.  
  2. Lê o **primeiro slice** de CT (`pydicom.dcmread`) e extrai:
     - `PatientName` e `PatientID`  
     - `StudyInstanceUID`, `SeriesInstanceUID`  
     - `FrameOfReferenceUID`  
  3. Em `self.dataset` (RTPLAN carregado):
     - Atualiza **PatientName** / **PatientID**.  
     - Atualiza **StudyInstanceUID** / **SeriesInstanceUID**.  
     - Gera um novo `SOPInstanceUID` (opcional).  
     - Atualiza **FrameOfReferenceUID** no RTPLAN.  
     - Atualiza toda a sequência **ReferencedStudySequence → ReferencedSeriesSequence** para apontar à nova série de CT.  
     - Atualiza **ReferencedFrameOfReferenceSequence** (toda a árvore de referência ao Frame of Reference).
  4. Recarrega a **árvore de tags** e a **visualização de MLC/Jaws**, refletindo imediatamente as alterações, **sem** solicitar ao usuário que grave o arquivo.  
  5. Exibe alerta “RTPLAN Atualizado” informando que, para persistir em disco, basta usar “Salvar Como…”.

Essa operação garante que, em memória, o RTPLAN carregado passe a “vincular” todas as suas referências de imagem ao novo CT Phantom gerado.

---

## Estrutura do Projeto

