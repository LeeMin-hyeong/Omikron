from pyloid_builder.pyinstaller import pyinstaller
from pyloid_builder.optimize import optimize
from pyloid.utils import get_platform
from pathlib import Path
import shutil


main_script = './src-pyloid/main.py'
name = 'main'
helper_script = './src-pyloid/update_helper.py'
helper_name = 'update-helper'
dist_path = './dist'
work_path = './build'


if get_platform() == 'windows':
	icon = './src-pyloid/icons/tdm_icon.ico'
elif get_platform() == 'macos':
	icon = './src-pyloid/icons/tdm_icon.png'
else:
	icon = './src-pyloid/icons/tdm_icon.png'

if get_platform() == 'windows':
    optimize_spec = './src-pyloid/build/windows_optimize.spec'
elif get_platform() == 'macos':
    optimize_spec = './src-pyloid/build/macos_optimize.spec'
else:
    optimize_spec = './src-pyloid/build/linux_optimize.spec'



if __name__ == '__main__':
	main_add_data = [
		'--add-data=./src-pyloid/icons/:./src-pyloid/icons/',
		'--add-data=./dist-front/:./dist-front/',
		'--add-data=./LICENSE:./',
	]
	if Path('./license.json').exists():
		main_add_data.append('--add-data=./license.json:./')

	pyinstaller(
		main_script,
		[
			f'--name={name}',
			f'--distpath={dist_path}',
			f'--workpath={work_path}',
			'--clean',
			'--noconfirm',
			'--onedir',
			# '--onefile',
			'--windowed',
			'--hidden-import=selenium.webdriver.chrome.webdriver',
			*main_add_data,
			f'--icon={icon}',
		],
	)
	pyinstaller(
		helper_script,
		[
			f'--name={helper_name}',
			f'--distpath={dist_path}',
			f'--workpath={work_path}',
			'--clean',
			'--noconfirm',
			'--onefile',
			'--windowed',
			f'--icon={icon}',
		]
	)
	# The transition archive needs both names. They intentionally contain the
	# same merged application; main.exe remains the legacy updater entry point.
	shutil.copy2(f'{dist_path}/{name}/{name}.exe', f'{dist_path}/tdm.exe')

	if get_platform() == 'windows':
		optimize(f'{dist_path}/{name}/_internal', optimize_spec)
	elif get_platform() == 'macos':
		optimize(f'{dist_path}/{name}.app', optimize_spec)
	else:
		optimize(f'{dist_path}/{name}/_internal', optimize_spec)
