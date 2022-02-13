# -*- coding: utf-8 -*-
"""
Created on Wed Feb  9 18:52:42 2022

@author: Stang
"""

import os
import json

from flask import Flask, render_template, request, jsonify, send_file

from src.jeopardy_from_template import JeopardyBuilder

import sys

app = Flask(__name__)
app.config['JSONIFY_PRETTYPRINT_REGULAR'] = False


_question_key = [
    f"clue_{i+1}_{j+1}" for i in range(10) for j in range(5)
]

""" _question_key = list(zip(_question_key, [
    f"answer_{i+1}_{j+1}" for i in range(10) for j in range(5)
])) """

_question_key.extend([
    "clue_f_j"
   # "answer_f_j"
])

_categories = [
    f"category_{i+1}" for i in range(10)
]

builder = JeopardyBuilder(
    template = "app/templates/jeopardy_template.pptx"
)

@app.route("/")
def home():
    return render_template("home.html")

@app.route("/build",methods=["POST"])
def build():
    if request.method == 'POST':
        questions = {
            _key : request.form[_key] for _key in _question_key
        }

        categories = {
            _key : request.form[_key] for _key in _categories
        }

        pptx_path = builder.build(categories, questions)
        
        return send_file(filename_or_fp = pptx_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)