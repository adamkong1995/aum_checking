def to_check(val):
    color_val = ''
    color_val = check_3_pct(val)
    color_val = check_missing(val)
    color_val = check_no_change(val)
    return 'background-color: %s' % color_val


def check_3_pct(val):
    color_val = ''
    if val > 0.03 or val < -0.03:
        color_val = 'yellow'
    return color_val

def check_missing(val):
    missing = object()
    if val is missing:
        color_val = 'green'
    return color_val

def check_no_change(val):
    if val is 0:
        color_val = 'blue'
    return color_val




