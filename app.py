from flask import Flask, request, jsonify
import asyncio
from playwright_utils import PW
import threading
from schedule_utils import start_scheduler

app = Flask(__name__)
pw = PW()


@app.route('/cyber', methods=['GET'])
async def init():
    result = await pw.init()
    return jsonify({
        'code': 900,
        'data': result
    })


@app.route('/cyber/gjcx', methods=['POST'])
async def gjcx():
    data = request.get_json()
    date_start = data.get('date_start')
    date_end = data.get('date_end')
    id_no = data.get('id_no')
    list_data = await pw.gjcx(date_start, date_end, id_no)
    return jsonify({
        'code': 900,
        'data': list_data
    })

# 其他路由类似实现，如 zzcx、gjcx_pl、zzcx_pl 等

if __name__ == '__main__':
    # 启动定时任务调度器
    scheduler_thread = threading.Thread(target=start_scheduler)
    scheduler_thread.daemon = True
    scheduler_thread.start()
    app.run(host='0.0.0.0', port=2325)
