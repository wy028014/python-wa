from flask import Flask, request, jsonify
from playwright_utils import PW
import asyncio

app = Flask(__name__)
pw = PW()


async def validate_and_execute(data, required_fields, query_method):
    if not isinstance(data, list):
        return jsonify({
            'code': 400,
            'message': '请求数据必须是数组类型'
        })
    tasks = []
    for item in data:
        for field in required_fields:
            if item.get(field) is None:
                return jsonify({
                    'code': 400,
                    'message': f'每个数组元素必须包含 {", ".join(required_fields)} 字段'
                })
        task = asyncio.create_task(query_method(**item))
        tasks.append(task)
    try:
        results = await asyncio.gather(*tasks)
        return jsonify({
            'code': 900,
            'data': results
        })
    except Exception as e:
        return jsonify({
            'code': 500,
            'message': f'执行查询时发生错误: {str(e)}'
        })


@app.route('/cyber/glcx', methods=['POST'])
async def glcx():
    data = request.get_json()
    required_fields = ['date_start', 'date_end', 'id_no']
    return await validate_and_execute(data, required_fields, pw.glcx)


@app.route('/cyber/zzcx', methods=['POST'])
async def zzcx():
    data = request.get_json()
    required_fields = ['train_date', 'train_code',
                       'from_station', 'to_station']
    return await validate_and_execute(data, required_fields, pw.zzcx)


@app.route('/cyber/plgjcx', methods=['POST'])
async def plgjcx():
    data = request.get_json()
    date_start = data.get('date_start')
    date_end = data.get('date_end')
    id_no_list = data.get('id_no_list')
    if date_start is None or date_end is None or not isinstance(id_no_list, list):
        return jsonify({
            'code': 400,
            'message': '请求数据必须包含 date_start、date_end 字段，且 id_no 必须是数组类型'
        })
    results = await pw.plgjcx(date_start, date_end, id_no_list)
    return jsonify({
        'code': 900,
        'data': results
    })

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=2325)
