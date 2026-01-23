from flask import Blueprint, render_template
from GaiZhangYe.core.basic.file_manager import get_file_manager

pages_bp = Blueprint('pages', __name__)


@pages_bp.route('/')
def index():
    return render_template('index.html')


@pages_bp.route('/page/prepare-stamp')
def prepare_stamp_page():
    folders = ["Nostamped_Word", "Nostamped_PDF", "Stamped_Pages", "Temp"]
    return render_template('pages/prepare_stamp.html', folders=folders, word_files=[])


@pages_bp.route('/page/stamp-overlay')
def stamp_overlay_page():
    return render_template('pages/stamp_overlay.html', word_files=[], image_files=[])


@pages_bp.route('/page/word-to-pdf')
def word_to_pdf_page():
    return render_template('pages/word_to_pdf.html')
