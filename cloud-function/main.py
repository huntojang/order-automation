import functions_framework
import requests
from flask import jsonify

@functions_framework.http
def alimtalk_proxy(request):
    """알리고 알림톡 API 프록시"""
    if request.method != 'POST':
        return jsonify({'error': 'POST only'}), 405

    data = request.form.to_dict() if request.form else request.json or {}

    try:
        resp = requests.post(
            'https://kakaoapi.aligo.in/akv10/alimtalk/send/',
            data=data,
            timeout=10
        )
        return jsonify(resp.json()), resp.status_code
    except Exception as e:
        return jsonify({'error': str(e)}), 500
