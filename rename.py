import os, sys
l= os.listdir(os.getcwd())
for k in l:
	new = k.split("file_")[-1]
	old = k
	os.rename(old,new)
