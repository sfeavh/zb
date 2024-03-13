# > nul 2>&1 || @echo off && cls
# > nul 2>&1 || python "%0" && goto :eof

import datetime
import pandas
import os

print("程序开始", flush=True)


def convert_time_to_decimal(time_str):
    time_parts = time_str.split(":")
    hours = int(time_parts[0])
    minutes = int(time_parts[1])
    seconds = int(time_parts[2])
    microseconds = int(time_parts[3])
    total_seconds = hours * 3600 + minutes * 60 + seconds + microseconds / 1000000
    return "%.1f" % (total_seconds / 3600 * 60)


folder_path = "F:\zb"
merged_df = pandas.DataFrame()
for file_name in os.listdir(folder_path):
    if file_name.endswith("xls"):
        print("正在读取：%s" % file_name, flush=True)
        df = pandas.read_excel(os.path.join(folder_path, file_name))
        file_name = file_name[:10]
        if len(file_name) > 6:
            file_name = datetime.datetime.strptime(file_name, "%Y-%m-%d").strftime("%Y/%m/%d")
        df.insert(0, "日期", file_name)


        def total_allocate(n):
            match n:
                case "回声嘹亮":
                    return "CCTV-15音乐频道"
                case "地理·中国":
                    return "CCTV-10科教频道"
                case "节目预告(文艺)":
                    return "CCTV-3综艺频道"
                case "文化十分":
                    return "CCTV-3综艺频道"
                case "民歌·中国":
                    return "CCTV-15音乐频道"
                case "艺术人生":
                    return "CCTV-3综艺频道"
                case "梨园闯关我挂帅":
                    return "CCTV-11戏曲频道"
                case "戏曲频道 时段导视":
                    return "CCTV-11戏曲频道"
                case "影视剧场":
                    return "CCTV-11戏曲频道"
                case "中华艺苑(西)":
                    return "CCTV-西语频道"
                case "这就是中国(西)":
                    return "CCTV-3综艺频道"
                case "中国音乐电视":
                    return "CCTV-15音乐频道"
                case "梨园周刊":
                    return "CCTV-11戏曲频道"
                case "回声嘹亮":
                    return "CCTV-15音乐频道"
                case "合唱先锋":
                    return "CCTV-15音乐频道"
                case "音乐公开课":
                    return "CCTV-15音乐频道"
                case "动物传奇":
                    return "CCTV-3综艺频道"
                case "动物世界":
                    return "CCTV-1综合频道"
                case "幸福账单":
                    return "CCTV-3综艺频道"
                case "全球中文音乐榜上榜":
                    return "CCTV-15音乐频道"
                case "音乐频道 时段导视":
                    return "CCTV-15音乐频道"
                case "2023秘境之眼":
                    return "CCTV-1综合频道"
                case "综艺频道 时段导视":
                    return "CCTV-3综艺频道"
                case "中央广播电视总台2023主持人大赛":
                    return "CCTV-1综合频道"
                case "2023年大国工匠年度人物发布仪式":
                    return "CCTV-1综合频道"
                case "2024总台戏曲频道元宵特别节目":
                    return "CCTV-11戏曲频道"
                case _:
                    return "不知道"


        df = df[~df['送播条数'].isin([0])]
        df["送播时长"] = df["送播时长"].apply(lambda x: convert_time_to_decimal(x))
        df["送播时长"] = df["送播时长"].astype('float')
        df["小时"] = df["送播时长"] / 60
        df["小时"] = df["小时"].round(1)
        df["名称"] = "高4"
        df["类型"] = "制作岛"
        df["位置"] = "现址"
        df["设备数量"] = 48
        df["频道"] = df["栏目名称"].map(total_allocate)

        merged_df = pandas.concat([merged_df, df[['名称', '类型', '位置', '日期', '设备数量', '送播条数', '频道', '栏目名称', '送播时长', '小时']]])


merged_df.to_excel(r'F:\zb\merged_data.xlsx', index=False)
print("完成", flush=True)
os.system('pause')