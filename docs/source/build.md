# Building

1. Clone the repository from https://github.com/lucsch/rustre

2. Create a virtual environnement::

        python3 -m venv venv
        source venv/bin/activate

3. Install required package  

        pip install -r requirements.txt`

  The required libraries are the following:

   - openpyxl
   - pytest
   - pytest-cov (code coverage)
   - coveralls (only needed to upload code coverage to coveralls.io)
   - sphinx (only needed for building the documentation)
   - myst-parser (markdown support for sphinx)
   - sphinx_rtd_theme (read the doc sphinx theme)

:::{Note}
If you do not plan to build the documentation, the last three libraries are not needed
:::

## Building the tests

      pytest test --log-cli-level=INFO

## Building the documentation

in the `RUSTRE/docs` folder run the following command:

      make html


