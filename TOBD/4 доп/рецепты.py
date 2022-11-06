import xlwings as xw
import numpy as np

# @xw.func is a decorator. 
# It must be added right before the def to let xlwings know this is a user-defined function.
@xw.func
def prod(ingridient):
    wb = xw.Book.caller()
    recipies = wb.sheets[0]
    prices = wb.sheets[1]
    
    data = recipies.used_range.options(np.array).value
    top_right_address = zip(*np.where(data == 'Доля'))
    bot_left_address = zip(*np.where(data == 'Средний физический расход ресурсов'))
    top_right_bot_left_address = zip(top_right_address,bot_left_address)
    recipe_ranges = [((top, left+1), (bottom+1, right+2)) for ((top, right),(bottom, left)) in top_right_bot_left_address]
    ranges = [recipies.range(*address) for address in recipe_ranges]
    
    count = 0
    
    for rec_range in ranges:
        
        ingridients = rec_range[2,5:-2]
        try:
            ingr_ind = ingridients.value.index(ingridient)
        except ValueError:
            continue
        for product in rec_range[4:-4,1:].rows:
            if product[0].value:
                if product[4+ingr_ind].value:
#                  
                    count+=1
    return count

@xw.func
def prod_names(ingridient):
    wb = xw.Book.caller()
    recipies = wb.sheets[0]
    prices = wb.sheets[1]
    
    data = recipies.used_range.options(np.array).value
    top_right_address = zip(*np.where(data == 'Доля'))
    bot_left_address = zip(*np.where(data == 'Средний физический расход ресурсов'))
    top_right_bot_left_address = zip(top_right_address,bot_left_address)
    recipe_ranges = [((top, left+1), (bottom+1, right+2)) for ((top, right),(bottom, left)) in top_right_bot_left_address]
    ranges = [recipies.range(*address) for address in recipe_ranges]
    
    prod_names = []
    
    for rec_range in ranges:
        
        ingridients = rec_range[2,5:-2]
        try:
            ingr_ind = ingridients.value.index(ingridient)
        except ValueError:
            continue
        for product in rec_range[4:-4,1:].rows:
            if product[0].value:
                if product[4+ingr_ind].value:
#                  
                    prod_names.append([product[0].value])
    return prod_names
import sqlite3

@xw.func
def prod_count_bySQL(ingridient):
    conn = sqlite3.connect("netCost.db")
    cursor = conn.cursor()
    cursor.execute(f"""
    SELECT count(distinct(products.name))
    FROM products INNER JOIN recipes
    ON products.id = recipes.'product id' INNER JOIN ingridients
    ON recipes.'ingridient id' = ingridients.id
    WHERE ingridients.ingridient='{ingridient}';
    """)
    return cursor.fetchall()[0][0]
    
@xw.func
def prod_names_by_SQL(ingridient):
    conn = sqlite3.connect("netCost.db")
    cursor = conn.cursor()
    cursor.execute(f"""
    SELECT distinct(products.name)
    FROM products INNER JOIN recipes
    ON products.id = recipes.'product id' INNER JOIN ingridients
    ON recipes.'ingridient id' = ingridients.id
    WHERE ingridients.ingridient='{ingridient}';
    """)
    return cursor.fetchall()
