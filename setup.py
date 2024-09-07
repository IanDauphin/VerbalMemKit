from setuptools import setup, find_packages

setup(
    name='semanticCorrectionFrench',            
    version='0.1',                # The initial release version
    packages=find_packages(),     # Automatically find all packages
    install_requires=[            # List any dependencies here
        'openpyxl'
    ],
    description='A description of your package.',
    author='James Donelle',
    url='https://github.com/jamedonelle/semanticCorrectionFrench', 
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',      # Minimum Python version required
)