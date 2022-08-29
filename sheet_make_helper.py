department_dictionary = dict()
department_name = 'cse'

if department_name not in department_dictionary.keys():
    department_dictionary[department_name] = dict()
    department_name[department_name]['list'] = []
    department_dictionary[department_name]['participated'] =0
    department_dictionary[department_name]['total'] = 0
else:
    department_dictionary[department_name]['list'].extend(class_list)
    department_dictionary[department_name]['total']+= len(class_list)

