""" This module is written for handling computations of DEMATEL and ANP"""
import numpy as np


def sub_matrices(dim_count, crit_count, t):
    """ divides the super matrix into it's constitutional matrices

    i is the dimensions count, j criteria count, k is starting point & l end
    point for slicing our super matrix rows while m & n do the same for columns.
    """
    t_list = []
    i = j = 0
    lis = []
    while i < int(dim_count):
        lis.append(j)
        j += int(crit_count[i])
        i += 1
    # lis is an accumulative list of criteria count e.g: [0, 2, 5, 6] used later
    # for slicing
    k = 0
    h = 1
    while k < int(dim_count):
        # list is sorted based on rows
        m = 0
        n = 1
        while m < int(dim_count):
            if n == int(dim_count):
                if h == int(dim_count):
                    t_list.append(t[lis[k]:, lis[m]:])
                else:
                    t_list.append(t[lis[k]:lis[h], lis[m]:])
            elif h == int(dim_count):
                t_list.append(t[lis[k]:, lis[m]:lis[n]])
            else:
                t_list.append(t[lis[k]:lis[h], lis[m]:lis[n]])
            # the idea behind all these conditions is that if h or n which are
            # our end points reach their final amount, slicing should be changed
            m += 1
            n += 1
        h += 1
        k += 1
    return t_list


def td_maker(dimension_count, td_v_list):
    """ creating Td from their list"""
    td_v = np.zeros((int(dimension_count), int(dimension_count)))
    row = g = 0
    while row < int(dimension_count):
        col = 0
        while col < int(dimension_count):
            td_v[row, col] = np.mean(td_v_list[g])
            col += 1
            g += 1
        row += 1
    return td_v


def normalizing(matrix):
    """ normalizing: dividing every item by it's row sum"""
    norm_mat = np.zeros((len(matrix), len(matrix[0])))
    for i in range(0, len(matrix)):
        for j in range(0, len(matrix[0])):
            norm_mat[i, j] = (matrix[i, j] / np.sum(matrix[i]))
    return norm_mat


def t_norm_making(td_list, dimension_count):
    """ creating super matrix from list of subset matrices

    the idea behind this function is to stack every i matrix in list (i being
    dimension count and therefore count of matrices in a row and column)
    horizontally and then vertically to form the super matrix from the list
    """
    lis2 = []
    i = 0
    w = 0
    while i < int(dimension_count):
        lis1 = []
        j = 0
        while j < int(dimension_count):
            lis1.append(td_list[w])
            j += 1
            w += 1
        row = np.hstack(lis1)
        lis2.append(row)
        i += 1
    col = np.vstack(lis2)
    return col


def weigh_super_list(un_sup_list, Td_norm_trans, dimension_count):
    """ multiplying sublist matrix by it's corresponding element in Td

    c is row, b is column & v is the matrix index
    """
    wg_super_list = []
    v, b, c = 0, 0, 0
    while v < int(dimension_count) * int(dimension_count):
        if b >= int(dimension_count):
            c += 1
            b = 0
        # row by row
        tmp = un_sup_list[v] * Td_norm_trans[c, b]
        wg_super_list.append(tmp)
        b += 1
        v += 1
    return wg_super_list
