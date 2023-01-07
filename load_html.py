from jinja2 import Environment, FileSystemLoader
from datetime import datetime
import sys, os


def load_context_to_template(context,template_name):
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    environment = Environment(loader=FileSystemLoader(os.path.join(application_path, "templates/")))
    template = environment.get_template(template_name)
    html = template.render(context)
    return html


if __name__ == "__main__":
    pass

