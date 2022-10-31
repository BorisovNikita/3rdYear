import numpy as np
import xlwings as xw
import sqlite3
import os

def table_creator(cursor):
    cursor.execute("""CREATE TABLE IF NOT EXISTS ingridients
                      (id INTEGER PRIMARY KEY, ingridient text, price real,
                      'measure unit' text,
                      CONSTRAINT unq_ind_1 UNIQUE(ingridient))
                   """)
    cursor.execute("""CREATE TABLE IF NOT EXISTS groups
                      (id INTEGER PRIMARY KEY, name text,
                      CONSTRAINT unq_ind_2 UNIQUE(name))
                   """)
    cursor.execute("""CREATE TABLE IF NOT EXISTS products
                      (id INTEGER PRIMARY KEY, name text, weight real,
                      'wholesale price' real, 'retail price' real,
                      fraction real, 'cost price' real, 'group id' integer,
                      CONSTRAINT unq_ind_3 UNIQUE(name, 'group id'))
                   """)
    cursor.execute("""CREATE TABLE IF NOT EXISTS recipes
                      ('product id' integer, 'ingridient id' integer,
                      consumption real,
                      CONSTRAINT unq_ind_4 UNIQUE('product id', 'ingridient id'))
                   """)
    cursor.execute("""CREATE TABLE IF NOT EXISTS averages
                      ('group id' integer, 'ingridient id' integer,
                      avrg real, 
                      CONSTRAINT unq_ind_5 UNIQUE('group id', 'ingridient id'))
                   """)

def updater():
    conn = sqlite3.connect("./netCost.db")
    cursor = conn.cursor()
    table_creator(cursor)
#     wb = xw.Book("рецепты.xlsx")
    wb = xw.Book.caller()
    recipies = wb.sheets[0]
    prices = wb.sheets[1]
    
    recipies.range((1,1)).value = os.getcwd()
    
    data = recipies.used_range.options(np.array).value
    top_right_address = zip(*np.where(data == 'Доля'))
    bot_left_address = zip(*np.where(data == 'Средний физический расход ресурсов'))
    top_right_bot_left_address = zip(top_right_address,bot_left_address)
    recipe_ranges = [((top, left+1), (bottom+1, right+2)) for ((top, right),(bottom, left)) in top_right_bot_left_address]
    ranges = [recipies.range(*address) for address in recipe_ranges]
    
        
    for id_, row in enumerate(
        prices.range((1,1)).expand().rows
    ):
        name, price, measure_unit = row.value
        row.color = None
        cursor.execute(f"""SELECT *
        FROM ingridients
        where (ingridient='{name}');""")
        finded = cursor.fetchall()
        if finded:
            finded=finded[0]
            if price!=finded[2]:
                row[1].color = "#008000" 
            if measure_unit!=finded[3]:
                row[2].color = "#008000"  
        else:
            row.color = "#008000"
            
        cursor.execute(f"""INSERT INTO ingridients
        (ingridient, price, 'measure unit') 
        VALUES('{name}', {price}, '{measure_unit}')
        ON CONFLICT(ingridient) DO UPDATE SET price=excluded.price,
        'measure unit'=excluded.'measure unit';""")
        conn.commit()
        
    for rec_range in ranges:
        ingridients = rec_range[2,5:-2]
        ingr_avrgs = rec_range[-1,5:-2]
        ingrs_id = []
#     print(ingridients)

        group_name = rec_range[0][0].value
        cursor.execute(f"""INSERT OR IGNORE INTO groups
        (name) 
        VALUES('{group_name}');""")
        conn.commit()
        cursor.execute(f"""SELECT id
        FROM groups
        WHERE name='{group_name}';""")
        group_id = cursor.fetchall()[0][0]


        for ingr, ingr_avrg in zip(ingridients,ingr_avrgs):
            if not ingr.value:
                continue

            cursor.execute(f"""SELECT id
            FROM ingridients
            WHERE ingridient='{ingr.value}';""")
            curr_ingr_id = cursor.fetchall()[0][0]
            ingrs_id.append(curr_ingr_id)

            ingr.color = None
            ingr_avrg.color = None
            cursor.execute(f"""SELECT *
            FROM averages
            where "group id"={group_id} AND "ingridient id"={curr_ingr_id};""")
            finded = cursor.fetchall()
            if finded:
                finded=finded[0]
#                 recipies.range((1,1)).value = abs(ingr_avrg.value - finded[2]) 
                if abs(ingr_avrg.value - finded[2]) > 0.000001:
                    ingr_avrg.color = "#008000" 
            else:
                ingr.color = "#008000" 
                ingr_avrg.color = "#008000" 


            cursor.execute(f"""INSERT INTO averages
            ('group id', 'ingridient id', avrg) 
            VALUES({group_id}, {curr_ingr_id}, {ingr_avrg.value})
            ON CONFLICT ("group id", "ingridient id")
            DO UPDATE SET avrg=excluded.avrg;""")
            conn.commit()
            
        for product in rec_range[4:-4,1:].rows:
            if product[0].value:
                pr_rng = product
                pr_rng.color = None
                product = product.value
                cursor.execute(f"""SELECT *
                FROM products
                where name='{product[0]}' AND "group id"={group_id};""")
                finded = cursor.fetchall()
                if finded:
                    finded = finded[0]
                    for mm in range(1,4):
                        if abs(product[mm]-finded[mm+1])>0.000001:
                            pr_rng[mm].color = "#008000" 
                else:
                    pr_rng.color = "#008000" 
                
                
                cursor.execute(f"""INSERT INTO products
                (name, weight,'wholesale price', 'retail price',
                fraction, 'cost price', 'group id')
                VALUES('{product[0]}', {product[1]}, {product[2]},
                {product[3]}, {product[-2]}, {product[-1]}, {group_id})
                ON CONFLICT ("name", "group id")
                DO UPDATE SET weight=excluded.weight,
                'wholesale price'=excluded.'wholesale price',
                'retail price'=excluded.'retail price',
                fraction=excluded.fraction,
                'cost price'=excluded.'cost price';""")
                conn.commit()

                cursor.execute(f"""SELECT id
                FROM products
                WHERE name='{product[0]}';""")
                product_id = cursor.fetchall()[0][0]

                for consumption, ingr_id in zip(pr_rng[4:-5],ingrs_id):
                    if consumption.value:

                        consumption.color = None
                        cursor.execute(f"""SELECT *
                        FROM recipes
                        where "product id"='{product_id}' AND "ingridient id"={ingr_id};""")
                        finded = cursor.fetchall()
                        if finded:
                            finded = finded[0]

                            if abs(consumption.value-finded[2])>0.000001:
                                consumption.color = "#008000" 
                        else:
                            consumption.color = "#008000" 


                        cursor.execute(f"""INSERT INTO recipes
                        ("product id", "ingridient id", consumption)
                        VALUES({product_id}, {ingr_id}, {consumption.value})
                        ON CONFLICT
                        DO UPDATE SET consumption=excluded.consumption;""")
                        conn.commit()
 
        
   
