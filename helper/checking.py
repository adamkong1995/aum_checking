def to_check(val):
    color_val = check_3_pct(val)
    return color_val


def check_3_pct(val):
    color_val = ''
    if val > 0.03 or val < -0.03:
        color_val = 'yellow'
    return 'background-color: %s' % color_val
