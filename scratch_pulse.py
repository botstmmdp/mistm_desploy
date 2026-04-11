
import os

html_files = [f for f in os.listdir('.') if f.endswith('.html') and f not in ('index.html', 'opciones.html', 'admin_panel')]

for f in html_files:
    try:
        with open(f, 'r', encoding='utf-8') as file:
            content = file.read()
            
        new_content = content.replace(
            '<a href="opciones.html" class="nav-btn">',
            '<a href="opciones.html" class="nav-btn nav-home-pulse">'
        )
        
        if new_content != content:
            with open(f, 'w', encoding='utf-8') as file:
                file.write(new_content)
            print(f'Added pulse to {f}')
    except Exception as e:
        print(f'Error: {e}')
