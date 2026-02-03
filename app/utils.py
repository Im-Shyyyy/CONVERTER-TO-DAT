def format_taxpayer_id(tax_id):
    tax_id = (tax_id or '').strip().replace('-', '')
    return f"{tax_id[:3]}-{tax_id[3:6]}-{tax_id[6:]}" if len(tax_id) == 9 else tax_id
