#!/usr/bin/env python
# -*- coding: utf-8 -*-

import numpy as np
from mpl_toolkits.mplot3d import Axes3D
import matplotlib.pyplot as plt
from matplotlib import cm


def rating(playCount, skipCount):
	#return 
	return (1-np.exp(-playCount/(0.1+skipCount))) * np.exp(-skipCount/playCount)

def main():
	maxPlay = 30
	maxSkip = 5
	
	thearray = np.zeros((maxPlay, maxSkip+1))

	for i in range(0, maxPlay):
		for j in range(0, maxSkip+1):
			thearray[i][j] = rating(float(i+1), float(j))

	result = thearray

	fig=plt.figure(figsize=(5, 5), dpi=150)
	ax1=fig.add_subplot(111, projection='3d')

	xlabels = np.array(range(0, maxSkip+1) )
	xpos = np.arange(xlabels.shape[0])
	
	ylabels = np.array(range(1, maxPlay+1) )
	ypos = np.arange(ylabels.shape[0])

	xposM, yposM = np.meshgrid(xpos, ypos, copy=False)

	zpos = result
	zpos = zpos.ravel()

	dx = 0.5
	dy = 0.5
	dz=zpos

	ax1.w_xaxis.set_ticks(xpos + dx/2.)
	ax1.w_xaxis.set_ticklabels(xlabels)

	ax1.w_yaxis.set_ticks(ypos + dy/2.)
	ax1.w_yaxis.set_ticklabels(ylabels)

	values = np.linspace(0.2, 1., xposM.ravel().shape[0])
	colors = cm.rainbow(values)
	ax1.bar3d(xposM.ravel(), yposM.ravel(), dz*0, dx, dy, dz, color=colors)
	plt.show()

if __name__ == '__main__':
	main()