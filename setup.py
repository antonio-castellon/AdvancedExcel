import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

# Load the dependencies list
with open('requirements.txt') as f:
    REQUIREMENTS = f.read().splitlines()

setuptools.setup(
    name="advance_excel",
    version="1.0.0",
    author="Antonio Castellon",
    author_email="antonio@castellon.ch",
    description="Advance Excel manipulator",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/antonio-castellon/AdvancedExcel",
    packages=setuptools.find_packages(where='src'),  # packages= find_packages(exclude='dist'),
	package_dir={'': 'src'},
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    # install_requires=REQUIREMENTS,
    python_requires='>=3.10',
)
