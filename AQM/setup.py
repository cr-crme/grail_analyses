from setuptools import setup

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="Grail analyses",
    packages=[
        ".",
    ],
    license="LICENSE",
    keywords=["Grail Analyses", "AQM"],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
    ],
    include_package_data=False,
    python_requires=">=3.7",
    zip_safe=False,
)