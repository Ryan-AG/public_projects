import PySimpleGUI as sg
import threading
from mk32_main import *

rules = '''MK32 Automation App:
    - Click 'Start' to begin processing the report.
    - When the report is finished, check the PDF before submitting it to FDS listener.'''

complete_list = ["Status:"]

layout = [
    [  # * Row 1
        sg.Text(
            rules,
            size=(60, 3)
        ),
    ],
    [  # * Row 2
        sg.Button(
            'Start',
            key='-START-'
        ),
        sg.Button(
            'Cancel'
        ),
    ],
    [  # * Row 3
        sg.Listbox(
            values=complete_list,
            enable_events=True,
            size=(68, 20),
            key='-LISTBOX-'
        )],
    [  # * Row 4
        sg.ProgressBar(
            max_value=100,
            key='-PROGRESS-',
            size=(45, 20))
    ]
]

window = sg.Window(title='MK32 Automation', layout=layout)

print(csm)

# def main():
#     t1 = threading.Thread(target=mk32_app.main)
#     while True:
#         event, values = window.read()
#         progress = 0
#         if event == sg.WIN_CLOSED or event == 'Cancel':
#             break
#         if event == "-START-":
#             t1.start()


# t2 = threading.Thread(target=main)

# if __name__ == '__main__':
#     t2.start()
