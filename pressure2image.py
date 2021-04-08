# 面圧センサデータから圧力分布画像データへ変換
import os			# ディレクトリアクセス
import csv			# CSVアクセス
import openpyxl		# エクセルファイル制御
from openpyxl.styles import PatternFill		# セルの色を変更する関数

# 解析対象ファイルペアリストを取得
# dir_l,dir_rに含まれるextensionを含む
def get_file_list(dir_l, dir_r, extension):
	files_l = []
	files_r = []
	file_list = []
	files_l = os.listdir(dir_l)									# ファイル取得
	files_r = os.listdir(dir_r)
	file_list = set(files_l) & set(files_r)						# 同じファイル名のリスト
	file_list = [f for f in file_list if '.'+extension in f]	# 指定拡張子のみを抽出
	file_list.sort()											# 破壊的ソート
	return file_list

# 圧力データリスト(圧力,時間)を取得
def get_pressure_list(file_puressure):
	press_list = []
	time_list = []
	with open(file_puressure, 'r') as file:
		header = next(csv.reader(file))		# ヘッダー行を空読み
		reader = csv.reader(file)			# データ行を読出し
		for line in reader:					# データ→リスト
			# line[0]=番号, line[1]=時間[sec], line[2]=カフ圧[mmHg]...
			press_list.append(float(line[2]))
			time_list.append(float(line[1]))

	return press_list, time_list

# 最高加圧情報(圧力,時間)を取得
def get_max_pressure(press_list, time_list):
	max_pres = 0.0
	max_pres_time = 0.0
	for i in range(1, len(press_list)-1):	# リストデータから最大値を取得
		if abs(press_list[i-1] - press_list[i]) < 1.0 and (press_list[i] - press_list[i+1]) < 1.0:	# ノイズ判定
			if max_pres <= press_list[i]:						# 最大値判定
				max_pres = press_list[i]
				max_pres_time = time_list[i]

	return max_pres,max_pres_time

# 規定圧力へ減圧するまでの時間を取得
def get_time_reduced_specified_pressure(pressure_list, time_list, spec_pres, max_pres):
	flg_reached_max_pressure = False
	time_reduced_specified_pressure = 0
	for i in range(1, len(pressure_list)-1):
		if abs(pressure_list[i-1] - pressure_list[i]) < 1.0 and abs(pressure_list[i] - pressure_list[i+1]) < 1.0:	# ノイズ判定
			if pressure_list[i] >= max_pres:												# 最大圧力超え判定
				flg_reached_max_pressure = True
			if flg_reached_max_pressure == True and pressure_list[i] <= spec_pres:			# 最大加圧超えた後に規定圧力まで減圧したか
				time_reduced_specified_pressure = time_list[i]
				break		# 取得できたので検索ループを脱出
	return time_reduced_specified_pressure

# 面圧センサデータの時間毎の平均圧力値を取得
def get_pressure_list_surface_sensor(file_sensor):
	avg_pres_list = []
	time_list = []
	sum_list = []
	with open(file_sensor, 'r') as file:
		for i in range(5):
			header = next(csv.reader(file))		# ヘッダー行を空読み
		reader = csv.reader(file)			# データ行を読出し
		for line in reader:					# データ→リスト
			# line[0]=時間, line[1]～line[256]:Elen0～Elem255の圧力
			total_press = 0
			for i in range(1,257):
				total_press = total_press + float(line[i])*100		# floatのままだと小数点演算で誤差が出るので整数にする
			avg_pres_list.append(total_press/256/100)
			time_list.append(float(line[0]))
	return avg_pres_list, time_list

# 面圧データより、最高加圧となった時間tmaxを取得
def get_time_max_pressure_surface_sensor(avg_sensor_pres_list, sensor_time_list):
	time_max_pressure_surface_sensor = 0
	max_pres = 0
	for i in range(1, len(avg_sensor_pres_list)-1):
		if abs(avg_sensor_pres_list[i-1] - avg_sensor_pres_list[i]) < 1.0 and abs(avg_sensor_pres_list[i] - avg_sensor_pres_list[i+1]) < 1.0:
			if max_pres <= avg_sensor_pres_list[i]:
				max_pres = avg_sensor_pres_list[i]
				time_max_pressure_surface_sensor = sensor_time_list[i]
	return time_max_pressure_surface_sensor

# tmaxからTp経過後の面圧センサデータを取得
def get_surface_sensor_data_target_press(file_sensor, time_max_pressure_surface_sensor, time_specified_pressure):
	surface_sensor_data_target_press = []
	with open(file_sensor, 'r') as file:
		for i in range(5):
			header = next(csv.reader(file))		# ヘッダー行を空読み
		reader = csv.reader(file)				# データ行を読出し
		for line in reader:						# データ→リスト
			# line[0]=時間, line[1]～line[256]:Elen0～Elem255の圧力
			if float(line[0]) >= time_max_pressure_surface_sensor + time_specified_pressure:
				print("面圧センサデータが指定圧力まで減圧した事を検出した時間:" + line[0])
				for i in range(1,257):
					surface_sensor_data_target_press.append(float(line[i]))
				break
	return surface_sensor_data_target_press

# 一次元配列を二次元配列へ変換
def convert_linear2matrix(array_linear):
	matrix = []
	for i in range(16):
		line = []
		for j in range(15, -1, -1):
			line.append(array_linear[i*16+j])
		matrix.append(line)
	return matrix

# 二次元配列の平滑化 対象点を近傍点と平均化する
def smooth_matrix(matrix_in, size):
	matrix_out = [[0 for x in range(16)] for y in range(16)]
	for y in range(int(0+(size-1)/2),int(16-(size-1)/2)):				# matrix[y][x]ループ
		for x in range(int(0+(size-1)/2),int(16-(size-1)/2)):
			for j in range(int(-(size-1)/2), int((size-1)/2+1)):		# 近傍点ループ
				for i in range(int(-(size-1)/2), int((size-1)/2+1)):
					matrix_out[y][x] += matrix_in[y+j][x+i]
			matrix_out[y][x] = round(matrix_out[y][x] / (size*size), 2)	# 小数点第3位以下四捨五入(銀行丸め)
	return matrix_out

# 圧力に応じて色づけしたエクセルファイルを作成する
def make_excel_sheet(matrix, workbook, sheet_name):
	color_code = [	'7030A0',	# 紫       0～
					'002060',	# 濃い青  30～
					'0070C0',	# 青      60～
					'00B0F0',	# 水色    90～
					'00B050',	# 緑     120～
					'92D050',	# 薄い緑 150～
					'FFFF00',	# 黄     180～
					'FFC000',	# 橙     210～
					'FF0000',	# 赤     240～
					'C00000' ]	# 濃い赤 270～
	num_sheet = len(workbook.sheetnames)
	workbook.create_sheet(index=num_sheet+1, title=sheet_name)
	ws = workbook.worksheets[num_sheet]
	for j in range(16):
		for i in range(16):
			ws.cell(i+1,j+1).value = matrix[j][i]
			if matrix[j][i] >= 300:
				fgColor = color_code[len(color_code)-1]
			elif matrix[j][i] <= 0:
				fgColor = 'FFFFFF'
			else:
				fgColor = color_code[int(matrix[j][i]/30)]
			ws.cell(i+1,j+1).fill = PatternFill(patternType='solid', fgColor=fgColor)	# 値に応じたセルの色にする

# 全平滑化データの圧力を平均した圧力分布データ作成
def get_average_matrix_list(matrix_list):
	matrix_avg = [[0 for x in range(16)] for y in range(16)]
	for matrix in matrix_list:
		for j in range(16):
			for i in range(16):
				matrix_avg[j][i] = matrix_avg[j][i] + matrix[j][i]

	for j in range(16):
		for i in range(16):
			matrix_avg[j][i] = round(matrix_avg[j][i] / len(matrix_list), 2)

	return matrix_avg

# メインルーチン
def main():
	print('面圧センサデータから圧力分布データファイルを作成')

	# ユーザー設定値入力
	print('検出圧力:')
	target_press = float(input())
	print('圧力平均するサイズ(NxNでNは奇数):')
	average_size = float(input())

	current_directory = os.getcwd()								# カレントディレクトリを取得
	dir_pressure = current_directory + '/' + "EG1データ"
	dir_sensor   = current_directory + '/' + "面圧データ"
	workbook = openpyxl.Workbook()								# 出力用エクセルファイル
	smooth_matrix_list = []										# 平滑化した圧力分布データのリスト

	# 解析対象ファイルペアを取得
	file_list = get_file_list(dir_pressure, dir_sensor, "csv")
	print('対象ファイル:', file_list)

	# ファイル数分処理する
	for file in file_list:
		print('----------------------------------------')
		print('解析ファイル:', file)
		# EG1データより、最高加圧までの時間Tmaxを取得
		pressure_list, time_list = get_pressure_list(dir_pressure + '/' + file)
		max_pressure, max_pressure_time = get_max_pressure(pressure_list, time_list)
		print("最大圧力:", max_pressure, ",  時間:", max_pressure_time)

		# 最大加圧後指定圧力までの減圧時間Tp(≡time_specified_pressure)を取得
		time_specified_pressure = get_time_reduced_specified_pressure(pressure_list, time_list, target_press, max_pressure) - max_pressure_time
		print("指定圧力まで減圧するまでの時間:", time_specified_pressure)

		# 面圧センサデータの時間毎の平均圧力値を取得
		avg_sensor_pres_list, sensor_time_list = get_pressure_list_surface_sensor(dir_sensor + '/' + file)

		# 面圧センサデータより、最高加圧となった時間tmaxを取得
		time_max_pressure_surface_sensor = get_time_max_pressure_surface_sensor(avg_sensor_pres_list, sensor_time_list)
		print("面圧センサが検出した最大圧力に達した時間:", time_max_pressure_surface_sensor)

		# tmaxからTp経過後の面圧センサデータを取得
		surface_sensor_data_target_press = get_surface_sensor_data_target_press(dir_sensor + '/' + file, time_max_pressure_surface_sensor, time_specified_pressure)
		# print(surface_sensor_data_target_press)

		# 面圧センサデータ(一次元配列)を二次元配列へ展開する
		matrix = convert_linear2matrix(surface_sensor_data_target_press)
		# print(matrix)

		# 面圧センサデータの各点のデータを、近傍の点より平均化する(3x3,5x5など)
		smoothed_matrix = smooth_matrix(matrix, average_size)
		# print(smoothed_matrix)
		smooth_matrix_list.append(smoothed_matrix)

		# 圧力に応じて色づけしたエクセルファイルを作成する
		make_excel_sheet(smoothed_matrix, workbook, file.replace('.csv',''))

	# 全平滑化データの圧力を平均した圧力分布データ作成してエクセルファイルへ追加
	average_matrix_list = get_average_matrix_list(smooth_matrix_list)
	make_excel_sheet(average_matrix_list, workbook, '平均')

	# エクセルファイル保存
	workbook.save(current_directory + '/output.xlsx')

	print("終わり")
	owari = input()

main()
