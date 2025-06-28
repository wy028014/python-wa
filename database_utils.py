import mysql.connector


class DB:
    def __init__(self):
        self.pool = mysql.connector.connect(
            host='10.3.32.233',
            port=3306,
            user='wy',
            password='Wy028014.',
            database='wy'
        )

    def query(self, sql, values=None):
        cursor = self.pool.cursor(dictionary=True)
        cursor.execute(sql, values)
        if sql.strip().lower().startswith('select'):
            result = cursor.fetchall()
        else:
            self.pool.commit()
            result = cursor.lastrowid
        cursor.close()
        return result

    def sel(self, ztrybh):
        sql = 'SELECT * FROM `在逃库` WHERE `在逃人员编号` = %s'
        return self.query(sql, (ztrybh,))

    def ins(self, param):
        sql = 'INSERT INTO `在逃库` (`在逃人员编号`, `在逃人员类型`, `姓名`, `性别`, `证件号码`, `户籍地行政区划`, `境内外去向`, `案件类别`, `简要案情`, `立案日期`, `立案单位`, `上网时间`, `案件数量`) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)'
        return self.query(sql, param)


db = DB()
