import requests
import json

class SharePointGraph:
    graph_api_base_url = "https://graph.microsoft.com/v1.0"

    def __init__(self, access_token):
        self.access_token = access_token
        self.headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

    def create_sharepoint_folder(self, folder_name, parent_item_id=None, drive_id=None, site_id=None, resolve_conflict="fail"):
        if drive_id and parent_item_id:
            url = f"{self.graph_api_base_url}/drives/{drive_id}/items/{parent_item_id}/children"
        elif site_id and parent_item_id:
            url = f"{self.graph_api_base_url}/sites/{site_id}/drive/items/{parent_item_id}/children"
        else:
            return False, "É necessário fornecer o drive_id ou site_id e o parent_item_id"

        payload = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": resolve_conflict
        }

        response = requests.post(url, json=payload, headers=self.headers)
        if response.status_code in [200,201]:
            return True, response.json()['id']
        else:
            return False, f"Erro ao criar pasta: {response.text}"

    def list_drive_items(self, drive_id, item_id=None):
        if item_id:
            endpoint = f"{self.graph_api_base_url}/drives/{drive_id}/items/{item_id}/children"
        else:
            endpoint = f"{self.graph_api_base_url}/drives/{drive_id}/root/children"

        response = requests.get(endpoint, headers=self.headers)

        if response.status_code in [200, 201]:
            items = response.json()["value"]
            json_response = json.loads(response.content)
            formatted_json = json.dumps(json_response, indent=2)
            results = []
            for item in items:
                results.append({
                    "id": item["id"],
                    "name": item["name"]
                })
            return True, results
        else:
            return False, f"Erro ao listar pastas: {response.text}"
    
    def upload_file(self, file_path, file_name, parent_item_id, drive_id=None, site_id=None):
        if drive_id:
            url = f"{self.graph_api_base_url}/drives/{drive_id}/items/{parent_item_id}:/{file_name}:/content"
        elif site_id:
            url = f"{self.graph_api_base_url}/sites/{site_id}/drive/items/{parent_item_id}:/{file_name}:/content"
        else:
            return False, "É necessário fornecer o drive_id ou o site_id"

        with open(file_path, "rb") as file:
            file_content = file.read()
            
        response = requests.put(url, data=file_content, headers=self.headers)
        if response.status_code in [200,201]:
            return True, response.json()['id']
        else:
            return False, f"Erro ao realizar upload do arquivo. {response.text}"

    def create_shareable_link(self, item_id, drive_id=None, site_id=None, link_type="view", scope="organization"):
        if drive_id:
            url = f"{self.graph_api_base_url}/drives/{drive_id}/items/{item_id}/createLink"
        elif site_id:
            url = f"{self.graph_api_base_url}/sites/{site_id}/drive/items/{item_id}/createLink"
        else:
            return False, "É necessário fornecer o drive_id ou o site_id"

        payload = {
            "type": link_type,
            "scope": scope
        }

        response = requests.post(url, data=json.dumps(payload), headers=self.headers)

        if response.status_code in [200,201,202]:
            return True, response.json()['link']['webUrl']
        else:
            return False, f"Erro ao criar link compartilhável. {response.text}"
    
    def delete_file_in_folder(self, folder_id, file_id, drive_id=None, site_id=None,):
        if drive_id:
            children_url = f"{self.graph_api_base_url}/drives/{drive_id}/items/{folder_id}/children"
        elif site_id:
            children_url = f"{self.graph_api_base_url}/sites/{site_id}/drive/items/{folder_id}/children"
        else:
            return False, "É necessário fornecer o drive_id ou o site_id"

        url = f"{self.graph_api_base_url}/drives/{drive_id}/items/{file_id}"
        response = requests.delete(url, headers=self.headers)
        if response.status_code in [200, 201, 204]:
            return True, 'Arquivo deletado.'
        else:
            return False, f"Erro ao deletar arquivo. {response.text}"
    
    def get_list_items(self, site_id, list_id, order_by=None, top=None, page_limit=None):
        if site_id and list_id:
            url = f"{self.graph_api_base_url}/sites/{site_id}/lists/{list_id}/items?expand=columns,items(expand=fields)"
        else:
            return False, "É necessário fornecer o site_id e o list_id"
        if order_by:
            url += f"&$orderby={order_by}"
        if top:
            url += f"&$top={top}"

        all_items = []
        page_count = 0

        while url:
            response = requests.get(url, headers=self.headers)
            if response.status_code == 200:
                data = response.json()
                items = data["value"]
                all_items.extend(items)

                url = data.get("@odata.nextLink", None)

                page_count += 1
                if page_limit and page_count >= int(page_limit):
                    break
            else:
                return False, f"Erro ao obter itens da lista. {response.text}"
        return True, all_items
