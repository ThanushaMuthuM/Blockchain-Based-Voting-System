import sys
import os

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT_DIR)

from flask import Flask, jsonify, request, send_from_directory, Response
from blockchain import Blockchain
import csv
import io

app = Flask(__name__)

blockchain = Blockchain()
ADMIN_PASSWORD = "admin123"