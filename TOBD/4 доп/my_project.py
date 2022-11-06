import xlwings as xw

# @xw.func is a decorator. 
# It must be added right before the def to let xlwings know this is a user-defined function.
@xw.func
def double_sum(x, y):
    """Returns twice the sum of the two arguments"""
    # The function must return something so the returned value can be passed into Excel
    return 2 * (x + y)
