from setuptools import setup

setup(
    name='docx2bb_web',
    version="0.20",
    license="GPLv3",
    author='Sinan Salman',
    author_email='sinan [dot] salman [at] gmail [dot] com',
    description=('docx2bb is a tool for converting test questions in MS-word (*.docx) file into a BlackBoard (text) import file'),
    url="https://bitbucket.org/sinansalman/docx2bb/",
    packages=['docx2bb_web'],
    include_package_data=True,
    install_requires=['flask','docx'],
    python_requires=">=3.5",
)
