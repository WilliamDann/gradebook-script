from xhtml2pdf import pisa
import numpy   as np


def loadTemplate():
    template = open('./template.html', 'r')
    templateText = template.read()
    template.close()

    return templateText

def table_row(name, mx, score):
    if np.isnan(score):
        return ""

    return """
        <tr class="border-2 border-grey-400">
            <td class="px-2 border-2 border-grey-400">{Name}</td>
            <td class="px-2 border-2 border-grey-400">{Max}</td>
            <td class="px-2 border-2 border-grey-400">{Score}</td>
        </tr>
    """.format_map({"Name": name, "Max": mx, "Score": score})

# Export a gradebook sheet into a set of pdf files
def export(xlsx, grading):
    # print(xlsx)

    students    = xlsx.iloc[2:, 0]
    assignments = xlsx.iloc[1, 2:8]
    weights     = grading.iloc[0]

    template = None
    try:
        template = loadTemplate()
    except FileNotFoundError:
        print("template.html needs to be in the same folder as the export script")
        print("if file is missing, find it here: https://gist.github.com/WilliamDann/b35c40ca96f0baa8b7eea2ae13960f41")
        return None
    
    i = 0
    for student in students:
        print(student)

        scores  = xlsx.iloc[2 + i, 2:8]
        overall = xlsx.iloc[2 + i, 11] 
        i += 1

        rows   = ""
        j      = 0
        for assignment in assignments:
            weight = weights.get(assignment)
            score  = scores.iloc[j]
            j += 1

            rows += table_row(assignment, weight, score)

        output = template.format_map({
            "Student_Name": student,
            "Overall_Grade": overall,
            "Grades": rows
        })
        with open("./out/" + student + ".pdf", 'wb') as f:
            pisa.CreatePDF(output, dest=f)