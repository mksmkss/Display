import textwrap
import budoux

# test = "We were exploring the sky when people looked up the sky and screamed a few hundred miles away."
test = "The quick brown fox jumps over the lazy dog."
# test = "夜に散歩をしていた時にふと横を見たら、茂みから見える都会がいつもより煌めいて見えたので写真に収めました。"

# googleの日本語parserを使う.これにより自然な改行ができるようになった
parser = budoux.load_default_japanese_parser()
_description = parser.parse(test)
description_list = []
long = 0
line = ""
max_len = 18
for k in _description:
    if (long + len(k)) <= max_len:
        line += k
        long += len(k)
    else:
        description_list.append(line)
        line = k
        long = len(k)
description_list.append(line)
print(description_list)

print(textwrap.wrap(test, 10))  #
print(textwrap.wrap(test, 15))  # 15 is the width of the line
print(textwrap.wrap(test, 20))  # 20 is the width of the line
