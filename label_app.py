import pandas as pd
import requests
import time
import argparse
from tqdm import tqdm
import sys

# --- Configuration ---
# API endpoints
GEOCODING_API_ENDPOINT = "https://maps.googleapis.com/maps/api/geocode/json"
ELEVATION_API_ENDPOINT = "https://maps.googleapis.com/maps/api/elevation/json"

def get_elevation(lat, lon, api_key):
    """
    Calls the Google Elevation API to get the altitude or an error message.
    """
    params = {'locations': f'{lat},{lon}', 'key': api_key}
    try:
        response = requests.get(ELEVATION_API_ENDPOINT, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
        if data['status'] == 'OK' and len(data['results']) > 0:
            return int(round(data['results'][0]['elevation'])) # Return as integer
        else:
             # Return the specific error message from Google
             return f"高度APIエラー: {data.get('error_message', data.get('status', 'Unknown Error'))}"
    except requests.exceptions.RequestException as e:
        return f"高度APIリクエストエラー: {e}"

def get_google_address_for_label(lat, lon, api_key):
    """
    Calls Google Geocoding API and returns the most suitable formatted address or an error message.
    """
    params = {'latlng': f'{lat},{lon}', 'key': api_key, 'language': 'ja'}
    try:
        response = requests.get(GEOCODING_API_ENDPOINT, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()
    except requests.exceptions.RequestException as e:
        return f"住所APIリクエストエラー: {e}"
        
    if data['status'] == 'OK' and len(data['results']) > 0:
        for result in data['results']:
            # Skip results that are just plus codes
            if 'plus_code' in result and result.get('types') == ['plus_code']:
                continue
            
            addr = result.get('formatted_address', '')
            # Clean up the address
            if addr.startswith('日本、'):
                addr = addr.replace('日本、', '', 1)
            # Remove postal code prefix
            space_pos = addr.find(' ')
            if space_pos != -1 and addr[:space_pos].replace('〒', '').replace('-', '').isdigit():
                return addr[space_pos+1:].strip()
            return addr.strip()
            
    elif data['status'] == 'ZERO_RESULTS':
        return "エラー: 住所が見つかりません"
    else:
        # Return the specific error message from Google
        return f"住所APIエラー: {data.get('error_message', data.get('status', 'Unknown Error'))}"

def create_label(row, lat_col, lon_col, date_col, method_col, collector_col):
    """
    Creates the final formatted label string from a DataFrame row.
    Handles potential error messages in address or elevation.
    """
    # Extract data from the row
    date = row.get(date_col, '')
    method = row.get(method_col, '')
    collector = row.get(collector_col, '')
    
    # --- FIX ---
    # Use the new, unique column names from the API results
    full_address = row.get('api_address', '') 
    elevation_val = row.get('api_elevation', '') 
    # --- END FIX ---

    lat = row.get(lat_col, '')
    lon = row.get(lon_col, '')
    
    # Format elevation
    elevation_str = ""
    if isinstance(elevation_val, (int, float)):
        elevation_str = f"GPS({elevation_val}m)"
    elif pd.notna(elevation_val) and str(elevation_val) != '':
        # If elevation_val is an error message, display it
        elevation_str = f"GPS(エラー: {elevation_val})"
    else:
        elevation_str = "GPS(高度取得失敗)"

    # Format address
    if not full_address or 'エラー' in str(full_address) or 'Error' in str(full_address):
        address_str = f"住所取得エラー: {full_address}"
    else:
        address_str = f"JAPAN: {full_address}"

    # Build the label string line by line
    label = f"{address_str}\n"
    label += f"{elevation_str}\n"
    
    line3_parts = []
    # Ensure all parts are strings before joining
    if pd.notna(date) and date != '': line3_parts.append(str(date))
    if pd.notna(method) and method != '': line3_parts.append(str(method))
    if pd.notna(collector) and collector != '': line3_parts.append(str(collector))
    
    if line3_parts:
        label += ". ".join(line3_parts)
        if not label.endswith('.'):
             label += "."
    
    label += f"\nN{lat}, E{lon}"
    
    return label

def main():
    """ Main function to run the script. """
    parser = argparse.ArgumentParser(description='CSVファイル内の緯度経度から住所と高度を取得し、最終的なラベル形式の文字列を生成します。')
    parser.add_argument('api_key', help='Google Maps APIキー (Geocoding APIとElevation APIが有効であること)。')
    parser.add_argument('input_csv', help='入力CSVファイルのパス。')
    # --- MODIFICATION ---
    parser.add_argument('output_csv', help='ラベル情報を追加した出力CSVファイル (.csv) のパス。')
    # --- END MODIFICATION ---
    parser.add_argument('--lat_col', default='latitude', help='緯度が含まれる列の名前 (デフォルト: latitude)。')
    parser.add_argument('--lon_col', default='longitude', help='経度が含まれる列の名前 (デフォルト: longitude)。')
    parser.add_argument('--date_col', default='採集年月日', help='日付が含まれる列の名前 (デフォルト: 採集年月日)。')
    parser.add_argument('--method_col', default='採集方法', help='採集方法が含まれる列の名前 (デフォルト: 採集方法)。')
    parser.add_argument('--collector_col', default='採集者名', help='採集者名が含まれる列の名前 (デフォルト: 採集者名)。')
    
    args = parser.parse_args()

    print(f"入力ファイル: {args.input_csv}")
    try:
        df = pd.read_csv(args.input_csv)
    except FileNotFoundError:
        print(f"エラー: 入力ファイル '{args.input_csv}' が見つかりません。")
        sys.exit(1)
    except Exception as e:
        print(f"入力ファイルの読み込みエラー: {e}")
        sys.exit(1)
        
    temp_results = []
    for index, row in tqdm(df.iterrows(), total=df.shape[0], desc="ジオコーディング処理中"):
        lat = row.get(args.lat_col)
        lon = row.get(args.lon_col)
        
        result = {}
        if pd.notna(lat) and pd.notna(lon):
            address = get_google_address_for_label(lat, lon, args.api_key)
            elevation = get_elevation(lat, lon, args.api_key)
            
            # --- FIX ---
            # Use new, unique column names to avoid conflict
            result['api_address'] = address
            result['api_elevation'] = elevation
            # --- END FIX ---
        else:
            result['api_address'] = '入力データなし'
            result['api_elevation'] = ''
            
        temp_results.append(result)
        time.sleep(0.05) # Rate limiting

    results_df = pd.DataFrame(temp_results)
    
    # Combine original data with new API data
    df_combined = pd.concat([df.reset_index(drop=True), results_df], axis=1)

    # Generate the final label column
    df_combined['label'] = df_combined.apply(
        lambda row: create_label(
            row, args.lat_col, args.lon_col, args.date_col, 
            args.method_col, args.collector_col
        ),
        axis=1
    )
    
    # --- MODIFICATION: Output to CSV ---
    try:
        # Save to CSV with UTF-8-SIG encoding for Excel compatibility
        df_combined.to_csv(args.output_csv, index=False, encoding='utf-8-sig')
        print(f"処理が完了しました。結果を '{args.output_csv}' に保存しました。")
    except Exception as e:
        print(f"\nCSVファイルへの書き出し中にエラーが発生しました: {e}")
    # --- END MODIFICATION ---

if __name__ == '__main__':
    main()

