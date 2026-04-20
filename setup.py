"""
PPEditer setup script
"""
from setuptools import setup, find_packages

setup(
    name="ppediter",
    version="1.0.0",
    description="Planning Proposal Editor - PPT/PPTX editor",
    author="Noh JinMoon",
    author_email="",
    license="MIT",
    packages=find_packages(),
    python_requires=">=3.7",
    install_requires=[
        "PyQt5>=5.15.0",
        "python-pptx>=0.6.21",
        "Pillow>=9.0.0",
        "lxml>=4.6.0",
    ],
    entry_points={
        "console_scripts": [
            "ppediter=main:main",
        ],
        "gui_scripts": [
            "PPEditer=main:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Topic :: Office/Business :: Office Suites",
    ],
)
