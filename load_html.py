from jinja2 import Environment, FileSystemLoader
from datetime import datetime

residents = [{"name":"tommy",
             "test_date": "01-05-2023",
            "lotNumber": "9023923",
            "expirationDate": "01/30/2023"
            }, 
            {"name":"tommy",
             "test_date": "01-05-2023",
            "lotNumber": "9023923",
            "expirationDate": "01/30/2023"
            }]
            

def load_context_to_template(context,template_name):
    environment = Environment(loader=FileSystemLoader("templates/"))
    template = environment.get_template(template_name)
    html = template.render(context)
    return html


if __name__ == "__main__":
    context = {"residents":residents}
    html = load_context_to_template(context, "ResidentTestingTemplate.html")
    print(html)


