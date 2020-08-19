from datetime import datetime, timedelta


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


def get_dates_between(start_date, end_date):
    delta = end_date - start_date  # as timedelta
    result = []
    for i in range(delta.days + 1):
        day = start_date + timedelta(days=i)
        result.append(day)
    return result


def check_if_service_dates_in_range(service_dates, start_date, end_date):
    # first list of the dates in range
    date_range = get_dates_between(start_date, end_date)

    result_indices = []
    source_mask = dict()
    for i in range(len(date_range)):
        source_mask[date_range[i]] = 0

    discrepancy = []
    for i in range(len(service_dates)):
        from_dos = service_dates[i][0]
        to_dos = service_dates[i][1]
        cur_range = get_dates_between(from_dos, to_dos)
        for cur_date in cur_range:
            if check_if_date_in_range(cur_date.strftime('%m/%d/%Y'), start_date.strftime('%m/%d/%Y'),
                                      end_date.strftime('%m/%d/%Y')):
                source_mask[cur_date] = 1
            elif cur_date not in discrepancy:
                discrepancy.append(cur_date)
                if i not in result_indices:
                    result_indices.append(i)
    missing_dates = []
    for i in range(len(date_range)):
        cur_date = date_range[i]
        if source_mask[cur_date] == 0:
            missing_dates.append(cur_date)

    missing_dates = parse_date_ranges(missing_dates)
    discrepancy = parse_date_ranges(discrepancy)

    return result_indices, missing_dates, discrepancy


def parse_date_ranges_into_list(dates):
    def group_consecutive(dates):
        dates_iter = iter(sorted(set(dates)))  # de-dup and sort

        run = [next(dates_iter)]
        for d in dates_iter:
            if (d.toordinal() - run[-1].toordinal()) == 1:  # consecutive?
                run.append(d)
            else:  # [start, end] of range else singleton
                yield [run[0], run[-1]] if len(run) > 1 else run[0]
                run = [d]
        yield [run[0], run[-1]] if len(run) > 1 else [run[0]]

    return list(group_consecutive(dates)) if dates else []

def parse_date_ranges(dates):
    result = []
    result_list = parse_date_ranges_into_list(dates)
    for i in range(len(result_list)):
        if len(result_list[i]) == 2:
            start_date = result_list[i][0]
            end_date = result_list[i][1]
            formatted_date = start_date.strftime('%m/%d/%Y') + ' - ' + end_date.strftime('%m/%d/%Y')
            result.append(formatted_date)
        else: # == 1
            date = result_list[i][0]
            formatted_date = date.strftime('%m/%d/%Y')
            result.append(formatted_date)
    result = '\n '.join(str(elem) for elem in result)
    return result
