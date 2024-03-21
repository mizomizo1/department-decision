from flask import Flask, request, render_template, send_file
import pandas as pd
import pulp
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "ファイルがありません", 400
    file = request.files['file']
    if file.filename == '':
        return "ファイルが選択されていません", 400
    if file:
        filepath = os.path.join('uploads', file.filename)
        file.save(filepath)
        # 問題の解決
        problem = DepartmentalMatchingProblem(filepath)
        result, assign = problem.solve()
        if result == pulp.LpStatusOptimal:
            # 解が見つかった場合、結果のデータフレームをExcelファイルに保存してダウンロードさせる
            assign_df = problem.to_dataframe(assign)
            output_path = os.path.join('output', 'assignment.xlsx')
            assign_df.to_excel(output_path)
            return send_file(output_path, as_attachment=True)
        else:
            return "最適解が見つかりませんでした", 500

if __name__ == '__main__':
    app.run(debug=True)

class DepartmentalMatchingProblem:
    def __init__(self, data_file):
        self.data_file = data_file
        self.df_students, self.df_ng, self.df_departments, self.department_caps = self.load_data()

    def load_data(self):
        df_students = pd.read_excel(self.data_file, sheet_name='Sheet1')
        df_ng = pd.read_excel(self.data_file, sheet_name='Sheet2', index_col=0)
        df_departments = pd.read_excel(self.data_file, sheet_name='Sheet3')
        department_caps = dict(zip(df_departments['配属先名'], df_departments['定員']))
        return df_students, df_ng, df_departments, department_caps

    def solve(self):
        num_students = len(self.df_students)
        num_departments = len(self.df_departments)
        assign = pulp.LpVariable.dicts('assign', ((i, j) for i in range(num_students) for j in range(num_departments)), cat='Binary')
        problem = pulp.LpProblem('DepartmentalMatchingProblem', pulp.LpMinimize)
        problem += pulp.lpSum([self.df_students.iloc[i, j+2] * assign[(i, j)] for i in range(num_students) for j in range(num_departments)])
        for i in range(num_students):
            problem += pulp.lpSum([assign[(i, j)] for j in range(num_departments)]) == 1
        for j in range(num_departments):
            problem += pulp.lpSum([assign[(i, j)] for i in range(num_students)]) <= self.department_caps[self.df_departments.iloc[j, 0]]
        for i in range(num_students):
            for j in range(num_departments):
                if self.df_ng.iloc[i, :].isin([self.df_students.iloc[k, 0] for k in range(num_students)]).any():
                    ng_list = list(self.df_ng.iloc[i, :].dropna())
                    for ng in ng_list:
                        if self.df_students.iloc[i, 0] != ng:
                            problem += assign[(i, j)] + assign[(self.df_students[self.df_students['学籍番号'] == ng].index[0], j)] <= 1
        result = problem.solve(pulp.PULP_CBC_CMD(msg=1, threads=4, timeLimit=100))
        return result, assign

    def create_assignment_df(self, assign):
        num_students = len(self.df_students)
        num_departments = len(self.df_departments)
        assign_df = pd.DataFrame(index=self.df_students['学籍番号'], columns=self.df_departments['配属先名'])
        for i in range(num_students):
            for j in range(num_departments):
                if assign[(i, j)].value() == 1:
                    assign_df.loc[self.df_students.iloc[i, 0], self.df_departments.iloc[j, 0]] = 1
                else:
                    assign_df.loc[self.df_students.iloc[i, 0], self.df_departments.iloc[j, 0]] = 0
        return assign_df

    def print_results(self, assign):
        num_students = len(self.df_students)
        num_departments = len(self.df_departments)
        for i in range(num_students):
            for j in range(num_departments):
                if assign[(i, j)].value() == 1:
                    print(f"{self.df_students.iloc[i, 0]}: {self.df_students.iloc[i, 1]} -> {self.df_departments.iloc[j, 0]}")
        assignments = {}
        for i in range(num_students):
            for j in range(num_departments):
                if assign[(i, j)].value() == 1:
                    if self.df_departments.iloc[j, 0] not in assignments:
                        assignments[self.df_departments.iloc[j, 0]] = []
                    assignments[self.df_departments.iloc[j, 0]].append(self.df_students.iloc[i, 1])
        for department in assignments:
            print(f"{department}:")
            for student in assignments[department]:
                print(f"  {student}")
        student_prefs = []
        for i in range(num_students):
            prefs = []
            for j in range(num_departments):
                if assign[(i, j)].value() == 1:
                    prefs.append(j)
            student_prefs.append(prefs)
        for i in range(num_students):
            prefs_str = ', '.join([f"第{n+1}希望に決定: {self.df_departments.iloc[p, 0]}" for n, p in enumerate(student_prefs[i])])
            print(f"{self.df_students.iloc[i, 0]}: {self.df_students.iloc[i, 1]} ({prefs_str})")

    def to_dataframe(self, assign):
        num_students = len(self.df_students)
        num_departments = len(self.df_departments)
        assign_df = pd.DataFrame(index=self.df_students['学籍番号'], columns=self.df_departments['配属先名'])
        for i in range(num_students):
            for j in range(num_departments):
                if assign[(i, j)].value() == 1:
                    assign_df.loc[self.df_students.iloc[i, 0], self.df_departments.iloc[j, 0]] = 1
                else:
                    assign_df.loc[self.df_students.iloc[i, 0], self.df_departments.iloc[j, 0]] = 0
        return assign_df 