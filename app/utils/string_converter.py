import re

def camel_to_snake(name):
    # Insert an underscore before capital letters that follow lowercase letters
    s1 = re.sub(r'(.)([A-Z][a-z]+)', r'\1_\2', name)
    # Insert an underscore before a capital letter that follows a lowercase letter or digit,
    # then convert the whole string to lowercase.
    return re.sub(r'([a-z0-9])([A-Z])', r'\1_\2', s1).lower()

def snake_to_camel(snake_str):
    components = snake_str.split('_')
    # Keep the first component as lowercase, capitalize the first letter of the remaining components
    return components[0] + ''.join(word.capitalize() for word in components[1:])

def snake_to_pascal(snake_str):
    # Define abbreviations that should be fully uppercased.
    abbreviations = {"id", "db", "wip", "erp", "plc"}
    
    parts = snake_str.split('_')
    pascal_parts = [
        part.upper() if part in abbreviations else part.capitalize()
        for part in parts
    ]
    return ''.join(pascal_parts)