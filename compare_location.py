import requests
from math import radians, cos, sin, asin, sqrt


# 利用高德地图api实现地址和经纬度的转换
def geocode(address):
    parameters = {'address': address, 'key': '75a8a958a6d332c42b860a50c5230e0a'}
    base = 'http://restapi.amap.com/v3/geocode/geo'
    response = requests.get(base, parameters)
    answer = response.json()

    if len(answer['geocodes']) < 1:
        print("ERROR：未识别地址[%s]" % (address))
        return False
    else:
        # print(address + "的经纬度：", answer['geocodes'][0]['location'])
        return answer['geocodes'][0]['location']


def geodistance(lng1, lat1, lng2, lat2):
    # 经度 longitude， 纬度 latitude 地球半径R=6371公里
    earth_R = 6371
    lng1, lat1, lng2, lat2 = map(radians, [lng1, lat1, lng2, lat2])
    dlon = lng2 - lng1
    dlat = lat2 - lat1
    dis = 2 * asin(sqrt(sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2)) * earth_R * 1000
    return dis


# 比较两个地址文本，通过转化为经纬度计算距离，确定是否为同一地址。返回值为1是、0不是、-1异常
# compare_location(address1, address2, precision)
# address1:字符串类型
# address2:字符串类型
# precision:数字类型，精度以米为单位
def compare_location(address1, address2, precision):
    loc1 = geocode(address1)
    loc2 = geocode(address2)
    if loc1 != False and loc2 != False:
        loc1 = loc1.split(',')
        loc2 = loc2.split(',')
        distance = geodistance(float(loc1[0]), float(loc1[1]), float(loc2[0]), float(loc2[1]))
        # print(loc1, loc2, distance)
        if distance < precision:
            return 1
        else:
            return 0
    else:
        return -1


def distance_2locations(address1, address2):
    loc1 = geocode(address1)
    loc2 = geocode(address2)
    if loc1 != False and loc2 != False:
        loc1 = loc1.split(',')
        loc2 = loc2.split(',')
        distance = geodistance(float(loc1[0]), float(loc1[1]), float(loc2[0]), float(loc2[1]))
        # print(loc1, loc2, distance)
        return distance
    else:
        return -1


if __name__ == '__main__':
    # address = input("请输入地址:")
    address1 = "湖南省长沙市芙蓉区司法警官职业技术学院"
    address2 = "湖南省长沙市天心区凌云东路与凌云路交叉口西南200米-长沙理工大学综合教学楼(西南门)"
    jw1 = geocode(address1)
    jw2 = geocode(address2)
    print("记录单地址[%s 的经度：%f 纬度：%f]" % (address1, float(jw1.split(',')[0]), float(jw1.split(',')[1])))
    print("签卡地址  [%s 的经度：%f 纬度：%f]" % (address2, float(jw2.split(',')[0]), float(jw2.split(',')[1])))
    print("系统判断两者相距：%s ,系统判断是否为同一结果：%s" % (
        distance_2locations(address1, address2), compare_location(address1, address2, 500)))
