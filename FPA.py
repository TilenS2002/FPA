import pandas as pd
import os

# Berem datoteko
excel_file = 'data.xlsx'
df = pd.read_excel(excel_file)
# število vrstic pred filtracijo log P > 1 in log P < -1
print(len(df))
df = df[(df['logFC'] >= float(1)) | (df['logFC'] <= float(-1))].reset_index(drop=True)
# število vrstic po filtraciji
print(len(df))

output_dir = 'Results'

# Množica besed lastnosti celic v posamičnih funkcijah
set_cellCompartment = set(['granule', 'extracellular matrix', 'ECM', 'membrane', 'cell surface', 'Golgi', 'endoplasmic reticulum', 'ER', 'lysosome', 'lysosomal', 'alpha granule', 'dense granule', 'exosome', 'cytosol', 'cytoplasm', 'cytoskeleton', 'extracellular', 'mitochondrial', 'mitochondria', 'nucleus', 'nucleoplasm', 'immunoglobulin', 'blood microparticle'])
set_molecularFunction = set(['peptidase', 'peptidase activity', 'chemokine', 'chemokine activity', 'heparanase', 'heparin', 'integrin', 'growth factor', 'antigen', 'immunoglobulin', 'immune', 'immune receptor', 'complement', 'adenylate kinase', 'ATP', 'actin', 'metalloendopeptidase activity', 'metalloendopeptidase', 'metal ion'])
set_biologicalProcess = set(['immune response', 'immune', 'MHC', 'inflammatory', 'monocyte', 'macrophage', 'vascular', 'blood', 'angiogenesis', 'coagulation', 'platelet aggregation', 'platelet activation', 'chemotaxis', 'cell proliferation', 'migration', 'cell division', 'nerve', 'nervous', 'neuron', 'microglia', 'signaling pathway', 'metabolic process', 'cytokine', 'chemokine', 'interleukin', 'extracellular matrix'])

table_cellCompartment = ['granule', 'extracellular matrix', 'membrane', 'Golgi', 'endoplasmic reticulum', 'lysosome', 'alpha granule', 'dense granule', 'exosome', 'cytosol', 'extracellular', 'mitochondria', 'nucleus', 'immunoglobulin', 'blood microparticle']
# pobrisu , 'ECM', 'cell surface', 'ER', 'lysosomal', 'cytoplasm', 'cytoskeleton''mitochondrial', , 'nucleoplasm'
table_molecularFunction = ['peptidase', 'chemokine', 'heparanase', 'integrin', 'growth factor', 'immunoglobulin', 'adenylate kinase', 'actin', 'metalloendopeptidase activity']
# pobrisu , 'peptidase activity', 'chemokine activity', 'heparin''antigen',, 'immune', 'immune receptor', 'complement', 'ATP', 'metalloendopeptidase', 'metal ion'
table_biologicalProcess = ['immune response', 'inflammatory', 'macrophage', 'vascular', 'platelet aggregation', 'chemotaxis', 'cell proliferation', 'nerve', 'signaling pathway', 'metabolic process', 'cytokine', 'extracellular matrix']
# pobrisu, 'immune', 'MHC', 'monocyte', 'blood', 'angiogenesis', 'coagulation', 'platelet activation', 'migration', 'cell division' 'nervous', 'neuron', 'microglia' 'chemokine', 'interleukin'

name_array = []

data_array = []

dict_cellCompartment = {}
dict_molecularFunction = {}
dict_biologicalProcess = {}

# Ustvari mapo za rezultate, če ne obstaja
# prečisti lastnosti in jih shrani v slovar
# vsaka funkcija posamično
os.makedirs(output_dir, exist_ok=True)
for column in df.columns:
    if column == 'Gene_Names_prim9ary':
        name_array = df[column]
    if column == 'GO_cellCompartment':
        for id, cell in enumerate(name_array):
            data_array = df[column][id]
            if isinstance(data_array, str):
                data_array = data_array.split(" ")
                if data_array.__contains__('extracellular'):
                    replace = data_array.index('extracellular')
                    if (data_array[replace+1] == 'matrix'):
                        data_array[replace] = 'extracellular matrix'
                        del data_array[replace+1]
                if data_array.__contains__('surface'):
                    replace = data_array.index('surface')
                    if (data_array[replace-1] == 'cell'):
                        data_array[replace] = 'cell surface'
                        del data_array[replace-1]
                if data_array.__contains__('endoplasmic'):
                    replace = data_array.index('endoplasmic')
                    if (data_array[replace+1] == 'reticulum'):
                        data_array[replace] = 'endoplasmic reticulum'
                        del data_array[replace+1]
                if data_array.__contains__('alpha'):
                    replace = data_array.index('alpha')
                    if (data_array[replace+1] == 'granule'):
                        data_array[replace] = 'alpha granule'
                        del data_array[replace+1]
                if data_array.__contains__('dense'):
                    replace = data_array.index('dense')
                    if (data_array[replace+1] == 'granule'):
                        data_array[replace] = 'dense granule'
                        del data_array[replace+1]
                if data_array.__contains__('blood'):
                    replace = data_array.index('blood')
                    if (data_array[replace+1] == 'microparticle'):
                        data_array[replace] = 'blood microparticle'
                        del data_array[replace+1]
            else:
                data_array = []
            dict_cellCompartment[cell] = list(set(data_array).intersection(set_cellCompartment))
    if column == 'GO_molecularFunction':
        for id, cell in enumerate(name_array):
            data_array = df[column][id]
            if isinstance(data_array, str):
                data_array = data_array.split(" ")
                if data_array.__contains__('peptidase'):
                    replace = data_array.index('peptidase')
                    if (data_array[replace+1] == 'activity'):
                        data_array[replace] = 'peptidase activity'
                        del data_array[replace+1]
                if data_array.__contains__('chemokine'):
                    replace = data_array.index('chemokine')
                    if (data_array[replace+1] == 'activity'):
                        data_array[replace] = 'chemokine activity'
                        del data_array[replace+1]
                if data_array.__contains__('growth'):
                    replace = data_array.index('growth')
                    if (data_array[replace+1] == 'factor'):
                        data_array[replace] = 'growth factor'
                        del data_array[replace+1]
                if data_array.__contains__('receptor'):
                    replace = data_array.index('receptor')
                    if (data_array[replace-1] == 'immune'):
                        data_array[replace] = 'immune receptor'
                        del data_array[replace-1]
                if data_array.__contains__('adenylate'):
                    replace = data_array.index('adenylate')
                    if (data_array[replace+1] == 'kinase'):
                        data_array[replace] = 'adenylate kinase'
                        del data_array[replace+1]
                if data_array.__contains__('metalloendopeptidase'):
                    replace = data_array.index('metalloendopeptidase')
                    if (data_array[replace+1] == 'activity'):
                        data_array[replace] = 'metalloendopeptidase activity'
                        del data_array[replace+1]
                if data_array.__contains__('metal'):
                    replace = data_array.index('metal')
                    if (data_array[replace+1] == 'ion'):
                        data_array[replace] = 'metal ion'
                        del data_array[replace+1]
            else:
                data_array = []
            dict_molecularFunction[cell] = list(set(data_array).intersection(set_molecularFunction))
    if column == 'GO_biologicalProcess':
        for id, cell in enumerate(name_array):
            data_array = df[column][id]
            if isinstance(data_array, str):
                data_array = data_array.split(" ")
                if data_array.__contains__('response'):
                    replace = data_array.index('response')
                    if (data_array[replace-1] == 'immune'):
                        data_array[replace] = 'immune response'
                        del data_array[replace-1]
                if data_array.__contains__('aggregation'):
                    replace = data_array.index('aggregation')
                    if (data_array[replace-1] == 'platelet'):
                        data_array[replace] = 'platelet aggregation'
                        del data_array[replace-1]
                if data_array.__contains__('activation'):
                    replace = data_array.index('activation')
                    if (data_array[replace-1] == 'platelet'):
                        data_array[replace] = 'platelet activation'
                        del data_array[replace-1]
                if data_array.__contains__('proliferation'):
                    replace = data_array.index('proliferation')
                    if (data_array[replace-1] == 'cell'):
                        data_array[replace] = 'cell proliferation'
                        del data_array[replace-1]
                if data_array.__contains__('division'):
                    replace = data_array.index('division')
                    if (data_array[replace-1] == 'cell'):
                        data_array[replace] = 'cell division'
                        del data_array[replace-1]
                if data_array.__contains__('signaling'):
                    replace = data_array.index('signaling')
                    if (data_array[replace+1] == 'pathway'):
                        data_array[replace] = 'signaling pathway'
                        del data_array[replace+1]
                if data_array.__contains__('metabolic'):
                    replace = data_array.index('metabolic')
                    if (data_array[replace+1] == 'process'):
                        data_array[replace] = 'metabolic process'
                        del data_array[replace+1]
                if data_array.__contains__('extracellular'):
                    replace = data_array.index('extracellular')
                    if (data_array[replace+1] == 'matrix'):
                        data_array[replace] = 'extracellular matrix'
                        del data_array[replace+1]
            else:
                data_array = []
            dict_biologicalProcess[cell] = list(set(data_array).intersection(set_biologicalProcess))

# Preslikava lastnosti z istim pomenom v izbrana imena stolpcev
property_synonyms = {
    'extracellular matrix': ['extracellular matrix', 'ECM'],
    'membrane': ['membrane', 'cell surface'],
    'endoplasmic reticulum': ['endoplasmic reticulum', 'ER'],
    'lysosome': ['lysosome', 'lysosomal'],
    'cytosol': ['cytosol', 'cytoplasm', 'cytoskeleton'],
    'mitochondria': ['mitochondria', 'mitochondrial'],
    'nucleus': ['nucleus', 'nucleoplasm'],
    'peptidase': ['peptidase', 'peptidase activity'],
    'chemokine': ['chemokine', 'chemokine activity'],
    'heparanase': ['heparanase', 'heparin'],
    'antigen': ['antigen', 'immunoglobulin', 'immune', 'immune receptor', 'complement'],
    'adenylate kinase': ['adenylate kinase', 'ATP'],
    'metalloendopeptidase activity': ['metalloendopeptidase activity', 'metalloendopeptidase', 'metal ion'],
    'immune response': ['immune response', 'immune', 'MHC'],
    'monocyte': ['monocyte', 'macrophage'],
    'vascular': ['vascular', 'blood', 'angiogenesis', 'coagulation'],
    'platelet aggregation': ['platelet aggregation', 'platelet activation'],
    'cell proliferation': ['cell proliferation', 'migration', 'cell division'],
    'nerve': ['nerve', 'nervous', 'neuron', 'microglia'],
    'cytokine': ['cytokine', 'chemokine', 'interleukin']
}

synonym_to_canonical = {}
for canonical, synonyms in property_synonyms.items():
    for synonym in synonyms:
        synonym_to_canonical[synonym] = canonical

# Excel - GO_cellCompartment markers
table = []
for gene in name_array:
    row = [gene]
    gene_properties = set()
    for prop in dict_cellCompartment.get(gene, []):
        canonical = synonym_to_canonical.get(prop, prop)
        gene_properties.add(canonical)
    for property in table_cellCompartment:
        row.append('X' if property in gene_properties else ' ')
    table.append(row)

df_markers = pd.DataFrame(table, columns=['Gene_Names_primary'] + table_cellCompartment)

df_markers.to_excel(os.path.join(output_dir, 'Results_GO_cellCompartment_markers.xlsx'), index=False)

# Excel - GO_molecularFunction markers
table = []
for gene in name_array:
    row = [gene]
    gene_properties = set()
    for prop in dict_molecularFunction.get(gene, []):
        canonical = synonym_to_canonical.get(prop, prop)
        gene_properties.add(canonical)
    for property in table_molecularFunction:
        row.append('X' if property in gene_properties else ' ')
    table.append(row)

df_markers = pd.DataFrame(table, columns=['Gene_Names_primary'] + table_molecularFunction)

df_markers.to_excel(os.path.join(output_dir, 'Results_GO_molecularFuntion_markers.xlsx'), index=False)

# Excel - GO_biologicalProcess markers
table = []
for gene in name_array:
    row = [gene]
    gene_properties = set()
    for prop in dict_biologicalProcess.get(gene, []):
        canonical = synonym_to_canonical.get(prop, prop)
        gene_properties.add(canonical)
    for property in table_biologicalProcess:
        row.append('X' if property in gene_properties else ' ')
    table.append(row)

df_markers = pd.DataFrame(table, columns=['Gene_Names_primary'] + table_biologicalProcess)

df_markers.to_excel(os.path.join(output_dir, 'Results_GO_biologicalProcess_markers.xlsx'), index=False)