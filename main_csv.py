import configparser
import csv


def check_conditions(row):
    answer_time = float(row[answer_time_index].rstrip("min")) * 60  # 答题时间 紧跟最后一个题目后面
    answer_data = row[1:max_questions + 1]  # 答题数据 第一个数据是序号要去掉

    # Rule 1: Answer time < 90 seconds
    if answer_time < min_answer_time:
        return "答题时间过小"

    # Rule 2: Majority answer count / total questions > 90%; no Two reverse questions
    if max_same_answer > 0:  # 有最大答案相似度判断时
        if fir_question_num == 0:  # 没有反向题判断时
            answer_no_req = answer_data
        else:
            answer_no_req = answer_data[0: fir_question_num - 1] + answer_data[fir_question_num: max_questions - 1]

        majority_count = max(answer_no_req.count('一般'), answer_no_req.count('非常符合'),
                             answer_no_req.count('比较符合'),
                             answer_no_req.count('非常不符合'), answer_no_req.count('非常不符合'))
        if majority_count / max_questions > max_same_answer:
            return "同一答案相似度过大"

    # Rule 3: Two reverse questions have same or similar options
    if fir_question_num > 0:
        if (answer_data[fir_question_num - 1] == answer_data[sec_question_num - 1]
            and answer_data[fir_question_num - 1] != '一般') \
                or (answer_data[fir_question_num - 1] == '比较符合' and answer_data[sec_question_num - 1] == '非常符合') \
                or (answer_data[fir_question_num - 1] == '非常不符合' and answer_data[sec_question_num - 1] == '比较不符合'):
            return "反向题答案相似"

    return None


if __name__ == '__main__':

    # Read settings from config.ini
    try:
        config = configparser.ConfigParser()
        config.read('config.ini')

        min_answer_time = float(config.get('Settings', 'min_answer_time'))  # 最小答题时间
        fir_question_num = int(config.get('Settings', 'reverse_question_1'))  # 第1个相反题的题号
        sec_question_num = int(config.get('Settings', 'reverse_question_2'))  # 第2个相反题的题号
        max_questions = int(config.get('Settings', 'max_questions'))  # 最大题目数
        max_same_answer = float(config.get('Settings', 'max_same_answer'))  # 同一答案相似度阈值
    except Exception as e:
        print("Failed to read config:", e)
        exit()

    # Read data from CSV file
    data = []
    encodings = ['utf-8', 'latin-1', 'gbk', 'gb2312', 'utf-16', 'utf-32', 'big5']  # Add more encodings if needed
    for encoding in encodings:
        try:
            with open('data.csv', 'r', encoding=encoding) as file:
                csv_reader = csv.reader(file)
                header = next(csv_reader)  # Store header row
                data = list(csv_reader)
            break
        except UnicodeDecodeError:
            continue
        except FileNotFoundError:
            print("File not found.")
            exit()
        except Exception as e:
            print("Failed to open the file:", e)
            exit()

    if not data:
        print("Failed to open the file with any of the specified encodings.")
        exit()

    # Find answer time index
    answer_time_index = None
    for i, label in enumerate(header):
        if '答题时间间隔' in label:
            answer_time_index = i
            break

    if answer_time_index is None:
        print("答题时间间隔列未找到")
        exit()

    # Find 无效说明 index
    invalid_index = None
    for i, label in enumerate(header):
        if '无效说明' in label:
            invalid_index = i
            break

    # Apply rules and update data
    for row in data:
        rule = check_conditions(row)
        if rule:
            if invalid_index is not None:
                row[invalid_index] = rule
            else:
                row.append(rule)
        else:
            row.append("")  # If no rule is satisfied, leave the cell empty

    # Write data back to CSV file
    try:
        with open('data.csv', 'w', newline='', encoding=encoding) as file:
            csv_writer = csv.writer(file)
            csv_writer.writerow(header)  # Write header row
            csv_writer.writerows(data)  # Write updated data
    except Exception as e:
        print("Failed to write the file:", e)
        exit()

    print("CSV file updated successfully.")
