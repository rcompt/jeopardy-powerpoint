# -*- coding: utf-8 -*-
"""
Created on Wed Feb  9 18:52:42 2022

@author: Stang
"""

import os
import json

from flask import Flask, render_template, request, jsonify
import logging

log = logging.getLogger("Jeopardy_log.txt")

app = Flask(__name__)
app.config['JSONIFY_PRETTYPRINT_REGULAR'] = False

#class NumpyArrayEncoder(JSONEncoder):
#    def default(self, obj):
#        if isinstance(obj, numpy.ndarray):
#            return obj.tolist()
#        return JSONEncoder.default(self, obj)

@app.route("/")
def home():
    return render_template("home.html")

@app.route("/predict",methods=["POST"])
def predict():
    if request.data:
        #req_data = request.get_json()
        #text = request.json.get('TEXT')

        return

if __name__ == "__main__":
    app.run(debug=True)
    log.info("App is running")