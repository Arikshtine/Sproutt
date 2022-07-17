import json
import sys
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse, parse_qs

from FactorCalculator import FactorCalculator
from HealthClassDecider import HealthClassDecider

hostName = "localhost"
serverPort = 8080


class MyRequestHandler(BaseHTTPRequestHandler):
    healthClassDecider = HealthClassDecider(sys.argv[1])
    factorCalculator = FactorCalculator(sys.argv[2])

    def do_GET(self):
        query_components = parse_qs(urlparse(self.path).query)

        term = int(query_components["term"][0]) if "term" in query_components else None
        coverage = int(query_components["coverage"][0]) if "coverage" in query_components else None
        age = int(query_components["age"][0]) if "age" in query_components else None
        height = query_components["height"][0] if "height" in query_components else None
        weight = int(query_components["weight"][0]) if "weight" in query_components else None

        if term is None or coverage is None or age is None or height is None or weight is None:
            self.send_response(code=400, message="One of the arguments is missing")
        else:
            health_class = self.healthClassDecider.calculate_health_class(height, weight)
            factor = self.factorCalculator.calculate_factor(health_class, term, age, coverage)
            if health_class == "Declined":
                self.send_response(code=400, message="Customer quote is declined because his health class is declined")
            elif factor == FactorCalculator.ERROR_CODE:
                self.send_response(code=400, message="This coverage is unsupported")
            else:
                price = round(coverage / 1000 * factor, 3)
                json_data = {"price": price,
                             "health - class": health_class,
                             "term": term,
                             "coverage": coverage}

                self.send_response(code=200, message='here is your token')
                self.send_header(keyword='Content-type', value='application/json')
                self.end_headers()
                self.wfile.write(json.dumps(json_data).encode('utf-8'))


if __name__ == "__main__":
    webServer = HTTPServer((hostName, serverPort), MyRequestHandler)
    webServer.serve_forever()
