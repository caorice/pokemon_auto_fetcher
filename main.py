import requests, argparse, openpyxl, regex as re
from urllib.parse import urlencode, quote_plus
from bs4 import BeautifulSoup
from typing import Dict, List, Tuple, Any
from pydash import omit_by, is_none, filter_

TITLE_KEY_MAP = {
    'input card name': 'search_content',
    'level': 'level',
    'highest': 'highest',
    'lowest': 'lowest',
    'average': 'average',
    'source': 'source',
}

def omit_none(value: Dict) -> Dict:
    return omit_by(value, is_none)

def get_price_in_string(price: str):
    match_result = re.search(r'(\d+\.?\d*)', price.replace(',', ''))
    if not match_result:
        return 0
    return float(match_result.group(1))

def is_in_after_strip(level:str, title: str):
    return re.sub(r'\s+', '', level).lower() in re.sub(r'\s+', '', title).lower()

def get_key_by_value(data: Dict, value: str):
    return list(data.keys())[list(data.values()).index(value)]

def get_dict_list_from_excel(file_path: str, title_key_map: Dict[str, str]) -> List[Dict[str, str]]:
    if not title_key_map:
        return []
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    result = []
    index_key_map = get_index_key_map(ws, title_key_map)
    for row in ws.iter_rows(min_row = 2):
        data_item = {}
        for index in index_key_map:
            if row[index].value:
                data_item[index_key_map[index]] = row[index].value
        result.append(data_item)
    return result

def set_dict_list_to_excel(file_path: str, dict_list: List[Dict[str, str]], title_key_map: Dict[str, str], is_edit: bool = True):
    if not title_key_map:
        return
    wb = openpyxl.load_workbook(file_path) if is_edit else openpyxl.Workbook()
    ws = wb.active
    index_key_map = get_index_key_map(ws, title_key_map) if is_edit else {index: item for index,item in enumerate(title_key_map.values())}
    for index, key in index_key_map.items():
        title = get_key_by_value(title_key_map, key)
        if title:
            ws.cell(row=1, column=index + 1).value = title
        for row_index, data_item in enumerate(dict_list):
            if data_item.get(key):
                ws.cell(row=row_index + 2, column= index + 1).value = data_item.get(key)
    wb.save(file_path)

def get_product_list_from_search_by_text(search_content: str, min_price: int = None, max_price: int = None) -> Tuple[str, List[Dict[str, str]]]:
    search_params = omit_none({
        '_nkw': search_content,
        '_udlo': min_price,
        '_udhi': max_price,
    })
    target_url = 'https://www.ebay.com/sch/i.html?{}'.format(
        urlencode(search_params, quote_via=quote_plus)
    )
    resp = requests.get(target_url, headers={
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36',
    })
    result = []
    soup = BeautifulSoup(resp.text, 'html.parser')
    for item in soup.select('li[data-viewport^="{\\"trackableId"] .s-item__wrapper .s-item__info'):
        price_tag = item.select_one('.s-item__price')
        title_tag = item.select_one('.s-item__title')
        if price_tag and title_tag:
            price = get_price_in_string(price_tag.text)
            if price:
                result.append({
                    'price': price,
                    'title': title_tag.text,
                })
    return target_url, result

def get_output_data_item(search_content: str, min_price: int = None, max_price: int = None, level: str = None) -> Dict[str, str]:
    url, result = get_product_list_from_search_by_text(search_content, min_price, max_price)
    result = filter_(result, lambda item: is_in_after_strip(level, item['title'])) if level else result
    if not result:
        return {
            'search_content': search_content,
            'level': level,
            'highest': None,
            'lowest': None,
            'average': None,
            'source': url,
        }
    else:
        return {
            'search_content': search_content,
            'level': level,
            'highest': max([x['price'] for x in result]),
            'lowest': min([x['price'] for x in result]),
            'average': sum([x['price'] for x in result]) / len(result),
            'source': url,
        }

'''
input_data_list:
[
    {
        search_content: 'charizard unlimited 4/102 cgc 7',
        level: 'CGC 7', # optional
    }
]
'''
def get_output_data_list(input_data_list: List[Dict[str, str]], min_price: int = None, max_price: int = None) -> List[Dict[str, str]]:
    result = []
    for input_data in input_data_list:
        if input_data.get('search_content'):
            result.append(get_output_data_item(input_data.get('search_content'), min_price, max_price, input_data.get('level')))
    return result

def get_index_key_map(ws: Any, title_key_map: Dict[str, str]):
    index_key_map = {}
    for index, cell in enumerate(ws[1]):
        matched_titles = filter_(title_key_map.keys(), lambda x: is_in_after_strip(cell.value, x))
        if matched_titles:
            title = max(matched_titles, key=len)
            index_key_map[index] = title_key_map[title]
    return index_key_map

def main():
    parser = argparse.ArgumentParser(
        prog = 'Pokemon auto fetcher',
        description = 'Input excel file and update newest price data into it.'
    )
    parser.add_argument('file', help='The excel file to input', type=str)
    parser.add_argument('--min', help='The min price to search', required=False, type=float)
    parser.add_argument('--max', help='The max price to search', required=False, type=float)
    parser.add_argument('--dump', '-d', help='Dump template file to edit', required=False, default=False, const=True, nargs='?')
    args_data = parser.parse_args()
    target_file: str = args_data.file
    if args_data.dump:
        output_file = target_file if re.search(r'.+\.xlsx?$', target_file, re.IGNORECASE) else '{}.xlsx'.format(target_file)
        set_dict_list_to_excel(output_file, [], TITLE_KEY_MAP, is_edit=False)
        print('Out put template file to {}'.format(output_file))
    else:
        dict_list = get_dict_list_from_excel(target_file, TITLE_KEY_MAP)
        if not dict_list:
            print('No any data to process.')
            return
        print('Read {} items from {}'.format(len(dict_list), target_file))
        print('Download data from ebay...')
        output_dict_list = get_output_data_list(dict_list, min_price = args_data.min, max_price = args_data.max)
        print('Download data from ebay done.')
        set_dict_list_to_excel(target_file, output_dict_list, TITLE_KEY_MAP)
        print('Write {} items to {}'.format(len(output_dict_list), target_file))
        print('Done.')

if __name__ == '__main__':
    main()