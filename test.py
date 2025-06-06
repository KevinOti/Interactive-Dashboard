tasks = []

def add_task(description, priority ='normal'):
    tasks.append({"description": description, "priority": priority})

def view_tasks():
    for task in tasks:
        print(f"{task['priority'].upper()}: {task['description']}")


def view_task_sorted():
    sorted_task = sorted(tasks, key=lambda task: task['priority'], reverse=False)
    print(sorted_task)

    


add_task('Refactor expense tracker', 'high')
add_task('automation script')
add_task("Improve Power Bi Filtering", 'urgent')

view_tasks()
view_task_sorted()
