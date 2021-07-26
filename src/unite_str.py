def unite_str(string:str) -> str:
    """统一字符串，增加匹配精度   （去空格 统一半角 小写）
    """
    return string.replace(' ', '').replace('（', '(').replace('）', ')').replace('：', ':').lower()
