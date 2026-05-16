import os, http.server, socketserver

PORT = int(os.environ.get('PORT', 8080))
Handler = http.server.SimpleHTTPRequestHandler
Handler.extensions_map['.html'] = 'text/html'

with socketserver.TCPServer(('', PORT), Handler) as httpd:
    print(f'Serving on port {PORT}')
    httpd.serve_forever()
