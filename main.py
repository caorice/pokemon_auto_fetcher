import requests, argparse, openpyxl, regex as re, json, os
from urllib.parse import urlencode, quote_plus, quote
from bs4 import BeautifulSoup
from typing import Dict, List, Tuple, Any, Union
from pydash import omit_by, is_none, filter_
from requests.adapters import HTTPAdapter
from urllib3.util import Retry

retry_strategy = Retry(
    total=3,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["HEAD", "GET", "OPTIONS"]
)

adapter = HTTPAdapter(max_retries=retry_strategy)
http = requests.Session()
http.mount("https://", adapter)
http.mount("http://", adapter)

TITLE_KEY_MAP = {
    'input card name': 'search_content',
    'level': 'level',
    'highest': 'highest',
    'lowest': 'lowest',
    'average': 'average',
    'number of result': 'count',
    'source': 'source',
}

USER_AGENT = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36'

PLATFORM_EBAY = 'ebay'
PLATFORM_130POINT = '130point'
PLATFORM_130POINT_ALL = '130point-all'

def get_proxy() -> Union[Dict, None]:
    if os.getenv('PROXY'):
        return {
            'http': os.getenv('PROXY'),
            'https': os.getenv('PROXY'),
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
            if data_item.get(key) is not None:
                ws.cell(row=row_index + 2, column= index + 1).value = data_item.get(key)
    wb.save(file_path)

def get_product_list_in_search_from_ebay(search_content: str, min_price: int = None, max_price: int = None) -> Tuple[str, List[Dict[str, str]]]:
    search_params = omit_none({
        '_nkw': search_content,
        '_udlo': min_price,
        '_udhi': max_price,
        'LH_Sold': 1,
        'LH_Complete': 1,
    })
    target_url = 'https://www.ebay.com/sch/i.html?{}'.format(
        urlencode(search_params, quote_via=quote_plus)
    )
    resp = requests.get(
        target_url,
        headers={
            'User-Agent': USER_AGENT,
        },
        proxies=get_proxy()
    )
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

def get_product_list_in_search_from_130point(search_content: str, min_price: int = None, max_price: int = None) -> Tuple[str, List[Dict[str, str]]]:
    search_params = {
        'query': search_content,
        'type': 2,
        'subcat': -1,
    }
    resp = requests.post(
        "https://130point.com/wp_pages/sales/getDataParse.php",
        headers={
            "User-Agent": USER_AGENT,
            "Content-Type": "application/x-www-form-urlencoded",
            'Cookie': 'session=56254937;'
        },
        data = urlencode(search_params, quote_via=quote).replace('%20', '%2B'),
        proxies=get_proxy()
    )
    if resp.status_code != 200:
        raise Exception('Get product list from 130point failed: {}, please retry.'.format(resp.status_code))
    try:
        body_str = resp.json().get('body')
    except Exception as e:
        return 'https://130point.com/sales/', []
    body_data = json.loads(body_str)
    result = []
    for item in body_data:
        price = get_price_in_string(item.get('price'))
        if price and (max_price is None or price <= max_price) and (min_price is None or price >= min_price):
            result.append({
                'price': price,
                'title': item.get('title'),
            })
    print(result)
    return 'https://130point.com/sales/', result

def get_product_list_in_search_from_130point_all(search_content: str, min_price: int = None, max_price: int = None) -> Tuple[str, List[Dict[str, str]]]:
    search_params = {
        'query': search_content,
        'sort': 'EndTimeSoonest',
        'tab_id': 3,
        'tz': 'Asia/Shanghai',
        'width': 965,
        'height': 763,
        'mp': 'all',
    }
    resp = requests.post(
        "https://130point.com/wp_pages/sales/getCards.php",
        headers={
            "User-Agent": USER_AGENT,
            "Content-Type": "application/x-www-form-urlencoded",
            'Cookie': 'session=56254937;',
            'x-requested-with': 'XMLHttpRequest',
            'referer': 'https://130point.com/cards/'
        },
        data = urlencode(search_params, quote_via=quote_plus),
        proxies=get_proxy()
    )
    if resp.status_code != 200:
        raise Exception('Get product list from 130point all tab failed: {}, please retry.'.format(resp.status_code))
    soup = BeautifulSoup(resp.text, 'html.parser')
    result = []
    for row in soup.select('tr#dRow'):
        price_str = row.attrs.get('data-price')
        title_element = row.select_one('span#titleText a')
        price = price_str and float(price_str)
        result.append({
            'price': price,
            'title': title_element and title_element.text,
        })
    return 'https://130point.com/cards/', result

def get_product_list_in_search(search_content: str, platform: str, min_price: int = None, max_price: int = None) -> Tuple[str, List[Dict[str, str]]]:
    if platform == PLATFORM_EBAY:
        return get_product_list_in_search_from_ebay(search_content, min_price, max_price)
    elif platform == PLATFORM_130POINT:
        return get_product_list_in_search_from_130point(search_content, min_price, max_price)
    elif platform == PLATFORM_130POINT_ALL:
        return get_product_list_in_search_from_130point_all(search_content, min_price, max_price)
    else:
        raise Exception('Platform not supported: {}'.format(platform))

def get_output_data_item(search_content: str, platform: str, min_price: int = None, max_price: int = None, level: str = None) -> Dict[str, str]:
    url, result = get_product_list_in_search(search_content, platform, min_price = max_price, max_price = max_price)
    result = filter_(result, lambda item: is_in_after_strip(level, item['title'])) if level else result
    if not result:
        return {
            'search_content': search_content,
            'level': level,
            'highest': None,
            'lowest': None,
            'average': None,
            'source': url,
            'count': 0
        }
    else:
        return {
            'search_content': search_content,
            'level': level,
            'highest': max([x['price'] for x in result]),
            'lowest': min([x['price'] for x in result]),
            'average': sum([x['price'] for x in result]) / len(result),
            'source': url,
            'count': len(result)
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
def get_output_data_list(input_data_list: List[Dict[str, str]], platform: str, min_price: int = None, max_price: int = None) -> List[Dict[str, str]]:
    result = []
    for index, input_data in enumerate(input_data_list):
        print('Processing {} / {}...'.format(index + 1, len(input_data_list)))
        if input_data.get('search_content'):
            result.append(get_output_data_item(input_data.get('search_content'), platform, min_price, max_price, input_data.get('level')))
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
    parser.add_argument('--platform', '-p', help='The platform to search', required=False, default=PLATFORM_EBAY, choices=[PLATFORM_EBAY, PLATFORM_130POINT, PLATFORM_130POINT_ALL], type=str)
    args_data = parser.parse_args()
    target_file: str = args_data.file
    platform: str = args_data.platform
    if args_data.dump:
        output_file = target_file if re.search(r'.+\.xlsx?$', target_file, re.IGNORECASE) else '{}.xlsx'.format(target_file)
        set_dict_list_to_excel(output_file, [], TITLE_KEY_MAP, is_edit=False)
        print('Output template file to {}'.format(output_file))
    else:
        dict_list = get_dict_list_from_excel(target_file, TITLE_KEY_MAP)
        if not dict_list:
            print('No any data to process.')
            return
        print('Read {} items from {}'.format(len(dict_list), target_file))
        print('Download data from {}...'.format(platform))
        output_dict_list = get_output_data_list(dict_list, platform, min_price = args_data.min, max_price = args_data.max)
        print('Download data from {} done.'.format(platform))
        set_dict_list_to_excel(target_file, output_dict_list, TITLE_KEY_MAP)
        print('Write {} items to {}'.format(len(output_dict_list), target_file))
        print('Done.')

if __name__ == '__main__':
    main()