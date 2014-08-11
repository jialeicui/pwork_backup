import MySQLdb
import sys

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

def get_same_id(lhs, rhs, third = None):
    global cur
    cur.execute('select con_3 from '+lhs)
    db1=set(cur.fetchall())
    cur.execute('select con_3 from '+rhs)
    db2=set(cur.fetchall())

    res = (db1 & db2)

    if third:
        cur.execute('select con_3 from '+third)
        db3=set(cur.fetchall())
        pass
    res = res & db3

    ret = []
    for one in res:
        ret.append(one[0])
        pass

    return ret
    pass

def add_to_table(ids):
    global cur
    global conn

    cur.execute('truncate table same_id')

    sql = 'INSERT INTO same_id (id, con_3) VALUES (NULL, %s)'
    params = [ids[i] for i in xrange(len(ids))]

    cur.executemany(sql, params)
    conn.commit()
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
    global cur

    f = open(file, 'w')
    # f = sys.stdout
    f.write(get_header(db) + '\n')

    add_to_table(ids)
    cur.execute('select *,count(distinct('+db+'.con_3)) from same_id,'+db+' where '+db+'.con_3 = same_id.con_3 group by same_id.con_3')
    res = cur.fetchall()

    for r in res:
        line = ''
        for one in r[2:]:
            if one == None:
                strr = ''
            else:
                strr = str(one)

            line = line + strr + ','

        f.write(line+'\n')

    f.close()
    pass

def main():
    connect()
    ids = get_same_id('data_table', 'data_table2', 'data_table3')
    export(ids, 'data_table', '1.csv')
    export(ids, 'data_table2', '2.csv')
    export(ids, 'data_table3', '3.csv')
    close()
    pass


if __name__ == '__main__':
    main()





