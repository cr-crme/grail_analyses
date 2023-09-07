from setuptools import setup

with open("README.md", "r") as fh:
    long_description = fh.read()

setup(
    name="Grail analyses",
    packages=[
        "grail_analyses",
        "grail_analyses/quantified_walking_analysis",
        "grail_analyses/game_analyses"
    ],
    license="LICENSE",
    keywords=["Grail Analyses", "Walking", "Quantified Walking Analysis"],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
    ],
    include_package_data=False,
    python_requires=">=3.10",
    zip_safe=False,
)