from os import path
from setuptools import setup, find_packages


here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.md'), encoding='utf-8') as readme_file:
    readme = readme_file.read()


with open(path.join(here, 'VERSION'), encoding='utf-8') as version_file:
    version = version_file.read()

with open(path.join(here, 'requirements.txt')) as requirements_file:
    # Parse requirements.txt, ignoring any commented-out lines.
    requirements = [line for line in requirements_file.read().splitlines()
                    if not line.startswith('#')]


setup(
    name="botcity_framework_t2c",
    version=version,
    description="Framework used by T2C to help develop automations in BotCity",
    long_description=readme,
    long_description_content_type='text/markdown',
    packages=find_packages(exclude=['docs', 'tests']),
    include_package_data=True,
    package_data={
        "framework": [
            # When adding files here, remember to update MANIFEST.in as well,
            # or else they will not be included in the distribution on PyPI!
            # 'path/to/data_file',
            'classes_t2c/T2CCloseAllApplications.py',
            'classes_t2c/T2CInitAllApplications.py',
            'classes_t2c/T2CInitAllSettings.py',
            'classes_t2c/T2CKillAllProcesses.py',
            'classes_t2c/T2CProcess.py',
            'classes_t2c/email/T2CSendEmail.py',
            'classes_t2c/email/T2CSendEmailOutlook.py',
            'classes_t2c/sqlite/T2CSqliteQueue.py',
            'classes_t2c/sqlserver/T2CSqlAnaliticoSintetico.py'
            'classes_t2c/relatorios/T2CRelatorios.py',
            'classes_t2c/utils/T2CExceptions.py',
            'classes_t2c/util/T2CMaestro.py',
            'resources',
        ]
    },
    install_requires=requirements,
)
