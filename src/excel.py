# -*- coding: utf-8 -*-

import pythoncom

def excel_range(n):
    return range(1, n+1)


COLLECT_METHOD = 1
COLLECT_ATTRIBUTE = 2

def collect(obj, rule):
    result = {}
    try:
        for key, value in rule.items():
            if value is None or callable(value):
                result[key] = getattr(obj, key)
                if callable(value):
                    result[key] = value(result[key])
            else:
                next_obj = getattr(obj, key)
                if isinstance(value, dict):
                    result[key] = collect(next_obj, value)
                else:
                    type, next_rule = value
                    count = 0
                    if type == COLLECT_METHOD:
                        count = next_obj().Count
                    elif type == COLLECT_ATTRIBUTE:
                        count = next_obj.Count
                    result[key] = [collect(next_obj(x), next_rule) for x in excel_range(count)]
    except (AttributeError, pythoncom.com_error):
        print(rule)
        raise
    return result
