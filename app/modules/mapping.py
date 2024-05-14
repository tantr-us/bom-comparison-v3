# Mapping from Agile numbers to customer numbers 
# base on a predefine value (dict)

# Sample setting
# mapping_setting = {
    # 'make': 'LFLIEP',
    # 'buy': 'LFLIE',
    # 'consigned suffix': 'CS',
    # 'customer docs': 'CUS',
    # 'rev delimiter': '/',
    # 'special delimiter': '-',
    # 'sample customer number': '8000-0201-000',
# }

# mapping_setting2 = {
    # 'make': 'LFCEPH',
    # 'buy': 'LFCEP',
    # 'consigned suffix': 'CS',
    # 'customer docs': 'CUS',
    # 'rev delimiter': '/',
    # 'special delimiter': None,
    # 'sample customer number': None,
# }

def transform_to_customer_number(part_number, mapping_setting):
    skip_prefix_list = ['TRA', 'MPI', 'BOM', 'DWG', 'FAB', 'GER', 'CAD', 'SCH', 'NET', 'TPI', 'RWK', 'PIP', 'FIX', 'FIA']

    make_prefix = mapping_setting['make']
    buy_prefix = mapping_setting['buy']
    cs_suffix = mapping_setting['consigned suffix']
    customer_doc_prefix = mapping_setting['customer docs']
    rev_delimiter = mapping_setting['rev delimiter']
    special_delimiter = mapping_setting['special delimiter']
    sample_cust_number = mapping_setting['sample customer number']
    
    # Set customer_number = part_number, the customer_number will go through the mapping process.
    # In the case the part_number has been manually mapped, the number will remain the same and
    # still be able to complete the compare process (ZOOX case)
    
    customer_number = part_number
    
    for skip_prefix in skip_prefix_list:
        if customer_number.startswith(skip_prefix):
            return customer_number
    
    # Determine if the part number is document or part:
    if customer_number.startswith(customer_doc_prefix):
        customer_number = customer_number.replace(customer_doc_prefix, '')
    else:
        # Remove prefix
        if customer_number.startswith(make_prefix):
            customer_number = customer_number.replace(make_prefix, '')
        elif customer_number.startswith(buy_prefix):
            customer_number = customer_number.replace(buy_prefix, '')
            
        if rev_delimiter in customer_number:
            customer_number = customer_number.split(rev_delimiter)[0]
        
        if customer_number.endswith(cs_suffix):
            customer_number = customer_number.replace(cs_suffix, '')
        
        if special_delimiter != None:
            pos_list = [pos for pos, char in enumerate(sample_cust_number) if char == special_delimiter]
            for pos in pos_list:
                customer_number = customer_number[:pos] + special_delimiter + customer_number[pos:] 
    return customer_number