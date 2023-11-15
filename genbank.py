from Bio import SeqIO
import pandas as pd
import glob

# 创建一个空的DataFrame，用于保存提取到的信息
df = pd.DataFrame(columns=['Locus', 'Definition', 'Accession', 'Version', 'Keyword', 'Source', 'Source_organism',
                           'Feature_source_organism', 'Feature_source_mol_type', 'Feature_source_strain',
                           'Feature_source_isolate', 'Feature_source_isolation_source', 'Feature_source_db_xerf',
                           'Feature_source_host', 'Feature_source_country', 'Feature_source_collection_date',
                           'Feature_source_collected'])

# 获取所有GenBank文件路径
genbank_files = glob.glob("path/to/genbank/files/*.gb")

# 遍历GenBank文件
for genbank_file in genbank_files:
    # 读取GenBank文件
    record = SeqIO.read(genbank_file, "genbank")

    # 提取信息
    locus = record.name
    definition = record.description
    accession = record.annotations['accessions'][0]
    version = record.annotations['version']
    keyword = record.annotations['keywords'][0] if 'keywords' in record.annotations else ''
    source = record.annotations['source']
    source_organism = record.annotations['organism']
    
    feature_source_organism = ''
    feature_source_mol_type = ''
    feature_source_strain = ''
    feature_source_isolate = ''
    feature_source_isolation_source = ''
    feature_source_db_xref = ''
    feature_source_host = ''
    feature_source_country = ''
    feature_source_collection_date = ''
    feature_source_collected = ''

    if 'references' in record.annotations:
        references = record.annotations['references']
        if references:
            feature_source_organism = references[0].journal_name
            feature_source_mol_type = references[0].title
            feature_source_strain = references[0].authors
            feature_source_isolate = references[0].journal
            feature_source_isolation_source = references[0].volume
            feature_source_db_xref = references[0].pages
            feature_source_host = references[0].pubmed_id
            feature_source_country = references[0].medline_id
            feature_source_collection_date = references[0].comment
            feature_source_collected = references[0].pubmed_id

    # 将提取到的信息添加到DataFrame中
    df = df.append({'Locus': locus, 'Definition': definition, 'Accession': accession, 'Version': version,
                    'Keyword': keyword, 'Source': source, 'Source_organism': source_organism,
                    'Feature_source_organism': feature_source_organism,
                    'Feature_source_mol_type': feature_source_mol_type,
                    'Feature_source_strain': feature_source_strain,
                    'Feature_source_isolate': feature_source_isolate,
                    'Feature_source_isolation_source': feature_source_isolation_source,
                    'Feature_source_db_xerf': feature_source_db_xref,
                    'Feature_source_host': feature_source_host,
                    'Feature_source_country': feature_source_country,
                    'Feature_source_collection_date': feature_source_collection_date,
                    'Feature_source_collected': feature_source_collected}, ignore_index=True)

# 将DataFrame导出到Excel文件
df.to_excel("output.xlsx", index=False)
