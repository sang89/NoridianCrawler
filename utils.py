from datetime import datetime


def convert_to_datetime_obj(date_time_str):
    return datetime.strptime(date_time_str, '%m/%d/%Y')


def check_if_date_in_range(date, start, end):
    start_dt = convert_to_datetime_obj(start)
    end_dt = convert_to_datetime_obj(end)
    date_dt = convert_to_datetime_obj(date)

    if start_dt <= date_dt <= end_dt:
        return True
    return False


def look_up_diagnostic_pointer_table(code, pointer_array):
    # the pointer_array has the following format
    # [['A', 'something', 'E', 'other'], ['B', 'cat', 'F', 'dog'],...]]
    for i in range(len(pointer_array)):
        current_array = pointer_array[i]
        if code in current_array:
            index = current_array.index(code)
            return current_array[index + 1]
    return ''


def look_up_reason_pointer_table(reasons, pointer_array):
    # the pointer_array has the following format
    # [['A', 'something'], ['B', 'dog']]
    # reasons maybe 'something dog'
    reason_array = reasons.split(' ')
    result_array = []
    for reason in reason_array:
        for i in range(len(pointer_array)):
            current_array = pointer_array[i]
            if reason in current_array:
                index = current_array.index(reason)
                result_array.append(current_array[index + 1])
                break
    result_str = ''
    for i in range(len(result_array)):
        result_str += '[' + str(i+1) + '] ' + result_array[i] + '\n'
    return result_str

def create_name_combinations(first_name, last_name):
    # want to get all combinations of type result = [[first1, last1], [first2, last2],...]
    result = []
    # 1: keep the same order
    result.append([first_name, last_name])
    # 2: reverse the order
    result.append([last_name, first_name])
    # 3: try to remove middle name in first name
    first_name_split = first_name.split(' ')
    if len(first_name_split) > 1:
        result.append([first_name_split[0], last_name])
        result.append([last_name, first_name_split[0]])
        result.append([last_name + ' ' + first_name_split[1], first_name_split[0]])
        result.append([first_name_split[0], last_name + ' ' + first_name_split[1]])
    # 4: try to remove middle name in last name
    last_name_split = last_name.split(' ')
    if len(last_name_split) > 1:
        result.append([last_name_split[0], first_name])
        result.append([first_name, last_name_split[0]])
        result.append([first_name + ' ' + last_name_split[1], last_name_split[0]])
        result.append([last_name_split[0], first_name + ' ' + last_name_split[1]])
    return result