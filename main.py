class Cell:
    def __init__(self):
        self.value = None

    def get_display_value(self, excel):
        if self.value == None: return ' '
        if self.value[0] != '=': return parse_float(self.value)
        formula = self.value[1:]
        final_value = self.formulate2(excel, formula)
        return final_value

    # def formulate1(self, excel, formula):
    #     if '+' in formula:
    #         terms = formula.split('+')
    #         return get_value_or_cell(excel, terms[0]) + get_value_or_cell(excel, terms[1])
    #     if '-' in formula:
    #         terms = formula.split('-')
    #         return get_value_or_cell(excel, terms[0]) - get_value_or_cell(excel, terms[1])
    #     if '*' in formula:
    #         terms = formula.split('*')
    #         return get_value_or_cell(excel, terms[0]) * get_value_or_cell(excel, terms[1])
    #     if '/' in formula:
    #         terms = formula.split('/')
    #         return get_value_or_cell(excel, terms[0]) / get_value_or_cell(excel, terms[1])
    #     return get_value_or_cell(excel, formula)

    def formulate2(self, excel, formula):
        tokens = parse_formula(Token(formula))
        # calculate token ex.['(', 'A1', '+', 'A2', ')', '/', '10.5']

        for i, value in enumerate(tokens):
            if is_cell(value):
                tokens[i] = get_value_or_cell(excel, value)

        return calculate_expression(tokens)

    def set_value(self, value):
        self.value = value

    def duplicate(self):
        new_cell = Cell()
        new_cell.value = self.value
        return new_cell

class Token:
    def __init__(self, string):
        self.string = string
        self.i = 0

    def empty(self):
        return self.i >= len(self.string)

    def peek(self):
        if self.empty(): return None
        return self.string[self.i]

    def next(self):
        n = self.peek()
        self.i += 1
        return n

def is_cell(s):
    if len(s) < 2: return False
    if not(ord('A') <= ord(s[0]) <= ord('Z')): return False
    for c in s[1:]:
        if not(ord('0') <= ord(c) <= ord('9')): return False
    return True

def is_op(s):
    return s in ['+', '-', '*', '/', '%']

def parse_formula(token):
    result = []
    limit = 100
    while not token.empty():
        first = token.peek()
        if first == ')':
            break
        elif first == '(':
            result.extend(parse_parenthesis(token))
        elif is_op(first):
            result.extend(parse_op(token))
        elif ord('A') <= ord(first) <= ord('Z'):
            result.extend(parse_cell(token))
        elif ord('0') <= ord(first) <= ord('9'): 
            result.extend(parse_number(token))

        limit -= 1
        if limit < 0: break

    return result

def parse_parenthesis(token):
    result = []
    result.append(token.next()) # (
    result.extend(parse_formula(token))
    result.append(token.next()) # )
    return result

def parse_number(token):
    nums = ''
    while not token.empty() and (ord('0') <= ord(token.peek()) <= ord('9') or token.peek() == '.'):
        nums += token.next()
    return [nums]

def parse_cell(token):
    cell = ''
    if ord('A') <= ord(token.peek()) <= ord('Z'):
        cell += token.next()
    while not token.empty() and (ord('0') <= ord(token.peek()) <= ord('9') or token.peek() == '.'):
        cell += token.next()
    return [cell]

def parse_op(token):
    return [token.next()]

def get_value_or_cell(excel, cell_or_value):
    value = parse_float(cell_or_value)
    if value != None: return value
    cell = cell_or_value
    value = get_excel(excel, cell).get_display_value(excel)
    return parse_float(value)

def parse_float(value, default = None):
    try:
        return float(value)
    except ValueError:
        return default

def display_excel(excel):
    print('--------- Excel ---------')
    for row in excel:
        for col in row:
            print(col.get_display_value(excel), end=' | ')
        print('')

def input_cell():
    cell = input('cell: ')
    value = input('value: ')
    return (cell, value)

def get_excel(excel, cell):
    i = int(cell[1:]) - 1
    j = ord(cell[0]) - ord('A')
    return excel[i][j]

def set_excel(excel, cell, value):
    i = int(cell[1:]) - 1
    j = ord(cell[0]) - ord('A')

    new_excel = []
    for row in excel:
        new_row = []
        for col in row:
            new_row.append(col.duplicate())
        new_excel.append(new_row)

    new_excel[i][j].set_value(value)
    return new_excel

def op_comperator(a, b):
    a_value = b_value = 0
    if a in ['+', '-']: a_value = 1
    elif a in ['*', '/', '%']: a_value = 2
    if b in ['+', '-']: b_value = 1
    elif b in ['*', '/', '%']: b_value = 2
    return a_value - b_value

def calculate_expression(tokens):
    postfixs = postfix_transformation(tokens)
    stack = []
    for token in postfixs:
        if is_op(token):
            a, b = stack[-2:]
            stack = stack[:-2]
            if token == '+': r = a + b
            if token == '-': r = a - b
            if token == '*': r = a * b
            if token == '/': r = a / b
            if token == '%': r = a % b
            stack.append(r)
        else:
            stack.append(token)
    return stack[-1]

def postfix_transformation(tokens):
    stack = []
    output = []
    for token in tokens:
        if parse_float(token) != None:
            output.append(parse_float(token))
        elif token == '(':
            stack.append(token)
        elif token == ')':
            while True:
                last = stack[-1]
                del stack[-1]
                if last == '(': break
                output.append(last)
        else: # op
            if len(stack) == 0 or stack[-1] == '(':
                stack.append(token)
            elif op_comperator(token, stack[-1]) == 1:
                stack.append(token)
            else: 
                while True:
                    last = stack[-1]
                    if not is_op(last): break
                    del stack[-1]
                    output.append(last)
                stack.append(token)

    while len(stack) > 0:
        last = stack[-1]
        del stack[-1]
        output.append(last)
    
    return output

width = 5
height = 4

excel = []

for i in range(height):
    row = []
    for j in range(width):
        row.append(Cell())
    excel.append(row)

display_excel(excel)

# while True:
#     cell, value = input_cell()
#     excel = set_excel(excel, cell, value)
#     display_excel(excel)

excel = set_excel(excel, 'A1', '1')
excel = set_excel(excel, 'A2', '2')
excel = set_excel(excel, 'A3', '=(A1+A2*8)/10.5')
excel = set_excel(excel, 'B1', '=((A3*100+(5)))')

display_excel(excel)
