setup_args={
        'name':'dataframe to autosize excel',
        'version':'1.0',
        'description':'Output pandas DataFrames into Excel Xlsx files with autofitted columns',
        'author':'Brennen Raimer',
        'url':'https://github.com/norweeg'
        }

try:
    from setuptools import setup, find_packages
except ImportError:
    from distutils.core import setup
    setup_args['packages'] = ["dataframe_to_autosize_excel"]
else:
    setup_args['packages'] = find_packages(exclude = ['contrib', 'docs', 'tests','reports','examples'])
    setup_args['project_urls'] = {'Source':'https://github.com/norweeg/DataFrame-to-Autofit-Xlsx'}
    setup_args['install_requires'] = ['pandas', 'xlsxwriter']
    setup_args['zip_safe'] = False
finally:
    setup(**setup_args)