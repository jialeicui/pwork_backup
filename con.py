import MySQLdb

cur = None
conn = None

def connect():
    global cur
    global conn
    try:
        conn=MySQLdb.connect(host='127.0.0.1',user='root',passwd='chinom',db='huirui',port=3306)
        cur=conn.cursor()
    except MySQLdb.Error,e:
        print "Mysql Error %d: %s" % (e.args[0], e.args[1])
    pass

def close():
    global cur
    global conn
    cur.close()
    conn.close()

def get_same_id(lhs, rhs):
    global cur
    cur.execute('select con_3 from '+lhs)
    db1=set(cur.fetchall())
    cur.execute('select con_3 from '+rhs)
    db2=set(cur.fetchall())
    res = (db1 & db2)

    ret = []
    for one in res:
        ret.append(one[0])
        pass

    return ret
    pass

def get_header(db):
    global cur
    cur.execute('show fields from ' + db);
    res = cur.fetchall()
    ret = ''
    for r in res:
        ret = ret + r[0] + ','
    return ret
    pass
def export(ids, db, file):
    f = open(file, 'w')
    f.write(get_header(db) + '\n')
    global cur
    for one in ids:
        q = 'select * from ' + db + ' where con_3=\'' + str(one) + '\''
        cur.execute(q);
        res = cur.fetchone()

        line = ''
        for r in res:
            if r == None:
                strr = ''
            else:
                strr = str(r)

            line = line + strr + ','

        f.write(line+'\n')
    f.close()
    pass

def main():
    connect()
    ids = get_same_id('data_table', 'data_table2')
    export(ids, 'data_table', '1.csv')
    export(ids, 'data_table2', '2.csv')
    close()
    pass


if __name__ == '__main__':
    main()





