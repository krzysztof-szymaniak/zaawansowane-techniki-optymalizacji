import RandomNumberGenerator as rng
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name as col

N = 5
SEED = 1


def init(seed):
    return rng.RandomNumberGenerator(seed)


def generate_costs(n: int, gen):
    array = []
    for _ in range(n):
        row = []
        for j in range(n):
            v = gen.nextInt(1, 50)
            row.append(v)
        array.append(row)
    return array


def write_array(array, ws, start=(0, 0)):
    for c, data in enumerate(array):
        ws.write_row(start[0] + c, start[1], data)


def generate_worksheet(costs, name):
    n = len(costs)
    workbook = xlsxwriter.Workbook(name)
    ws = workbook.add_worksheet()

    ws.write(0, 0, "Costs")
    tasks = [[f'Task {i + 1}' for i in range(n)]]
    workers = [[f'Worker {i + 1}'] for i in range(n)]
    write_array(tasks, ws, start=(0, 1))
    write_array(workers, ws, start=(1, 0))
    write_array(costs, ws, start=(1, 1))

    r = [0] * n
    decision_vars = [r] * n
    second_start_row = n + 3
    ws.write(second_start_row, 0, "Assigment")
    write_array(tasks, ws, start=(second_start_row, 1))
    write_array(workers, ws, start=(second_start_row + 1, 0))
    write_array(decision_vars, ws, start=(second_start_row + 1, 1))

    ws.write(second_start_row, n + 2, "Tasks assigned")
    for i in range(n):
        sum_start_row = second_start_row + 2 + i
        sum_start_col = col(1)
        sum_end_row = sum_start_row
        sum_end_col = col(n)
        formula = f'=SUM({sum_start_col}{sum_start_row}:{sum_end_col}{sum_end_row})'
        ws.write(second_start_row + 1 + i, n + 2, formula)

    ws.write(n + 3, n + 3, "Tasks constraint")
    ws.write_column(n + 4, n + 3, [1] * n)

    ws.write(second_start_row + n + 2, 0, "Total assigned")
    for i in range(n):
        sum_start_row = second_start_row + 2
        sum_start_col = col(i + 1)
        sum_end_row = sum_start_row + n
        sum_end_col = sum_start_col
        formula = f'=SUM({sum_start_col}{sum_start_row}:{sum_end_col}{sum_end_row})'
        ws.write(second_start_row + n + 2, i + 1, formula)

    ws.write(second_start_row + n + 3, 0, "People constraint")
    ws.write_row(second_start_row + n + 3, 1, [1] * n)
    ws.write(second_start_row + n + 3, n + 2, "Total cost")

    formula = f'=SUMPRODUCT(({col(1)}{2}:{col(n)}{n + 1})*({col(1)}{second_start_row + 2}:{col(n)}{second_start_row + n + 1}))'
    ws.write(second_start_row + n + 3, n + 3, formula)
    workbook.close()


def main():
    gen = init(SEED)
    array = generate_costs(N, gen)
    generate_worksheet(array, 'task_assignment.xlsx')


if __name__ == '__main__':
    main()
