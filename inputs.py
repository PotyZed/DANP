""" This module is written for capturing user inputs"""


def exp_count():
    """ this function retrieves the number of experts who answered the
     questionnaire"""
    experts_count = input('please enter how many applicants have filled your'
                          ' questionnaire?\n')
    while True:
        try:
            experts_count = int(experts_count)
        except ValueError:
            experts_count = input('please enter an integer number')
        else:
            return experts_count


def dim_count():
    """ these functions receive the count of dimensions and it's criteria"""
    dimension_count = input('how many dimensions does your research have?\n')
    return dimension_count


def crit_count(dimension_count):
    # To Do: could use some organizing (lots of recurring commands)
    criteria_count = []
    j = 0
    while j < int(dimension_count):
        if j == 0:
            while True:
                try:
                    first_dim_count = int(input("how many criteria does your "
                                                "1st dimension contain?\n"))
                except ValueError:
                    first_dim_count = input('please enter an integer number\n')
                else:
                    criteria_count.append(first_dim_count)
                    break
        elif j == 1:
            while True:
                try:
                    second_dim_count = int(input("how many criteria does your"
                                                 " 2nd dimension contain?\n"))
                except ValueError:
                    second_dim_count = input('please enter an integer number\n')
                else:
                    criteria_count.append(second_dim_count)
                    break
        elif j == 2:
            while True:
                try:
                    third_dim_count = int(input("how many criteria does your "
                                                "3rd dimension contain?\n"))
                except ValueError:
                    third_dim_count = input('please enter an integer number\n')
                else:
                    criteria_count.append(third_dim_count)
                    break
        else:
            while True:
                try:
                    other_count = int(input(f"how many criteria does your "
                                            f"{j + 1}th dimension contain?\n"))
                except ValueError:
                    other_count = input('please enter an integer number\n')
                else:
                    criteria_count.append(other_count)
                    break
        j += 1
    return criteria_count


def crit_name(dimension_count, criteria_count):
    """ naming criteria based on their dimension"""
    criteria_names = []
    d = 0
    while d < int(dimension_count):
        c = 0
        while c < int(criteria_count[d]):
            criteria_names.append(f'c{d + 1}{c + 1}')
            c += 1
        d += 1
    return criteria_names


def dim_name(dimension_count):
    """ naming dimensions"""
    dimension_names = []
    d = 0
    while d < int(dimension_count):
        dimension_names.append(f'd{d + 1}')
        d += 1
    return dimension_names
