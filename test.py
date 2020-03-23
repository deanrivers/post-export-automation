s = "000134_03062000294_IC-CHECK_PAM_DM.TXT"

start = s.find("03062000294_") + len('03062000294_')
end = s.find("_DM")
substring = s[start:end]
print(substring)

