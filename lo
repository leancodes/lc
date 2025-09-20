import numpy as np, matplotlib.pyplot as plt

def unitvec(theta):
    return np.array([np.cos(theta), np.sin(theta)])

R = 1.0
thetas = np.deg2rad([90, 210, 330])
rv = R / (1 + np.sqrt(3))
V = np.stack([rv*unitvec(t) for t in thetas])

t = np.linspace(0, 2*np.pi, 800)
xc, yc = R*np.cos(t), R*np.sin(t)

def angle(p, c): 
    return np.arctan2(p[1]-c[1], p[0]-c[0])
arcs_x, arcs_y = [], []
for i in range(3):
    ci = V[i]
    a = V[(i+1)%3]
    b = V[(i+2)%3]
    ra = angle(a, ci)
    rb = angle(b, ci)
    # major arc from a to b about center ci
    d = (rb - ra) % (2*np.pi)
    if d < np.pi:  # switch to major
        d = 2*np.pi - d
        ts = np.linspace(rb, rb + d, 400)
    else:
        ts = np.linspace(ra, ra + d, 400)
    r_arc = np.linalg.norm(a-ci)
    arcs_x.append(ci[0] + r_arc*np.cos(ts))
    arcs_y.append(ci[1] + r_arc*np.sin(ts))
xr = np.concatenate(arcs_x); yr = np.concatenate(arcs_y)

Rparam = R/3.0
td = np.linspace(0, 2*np.pi, 1200)
xd = 2*Rparam*np.cos(td) + Rparam*np.cos(2*td)
yd = 2*Rparam*np.sin(td) - Rparam*np.sin(2*td)
phi = thetas[0]
c, s = np.cos(phi), np.sin(phi)
xd_rot = c*xd - s*yd
yd_rot = s*xd + c*yd

plt.figure(figsize=(6,6))
plt.plot(xc, yc, 'k', linewidth=2)
plt.plot(xr, yr, 'b', linewidth=2)
plt.plot(xd_rot, yd_rot, 'r', linewidth=2)
plt.scatter(R*np.cos(thetas), R*np.sin(thetas), color='k', zorder=5)
plt.gca().set_aspect('equal', 'box')
plt.axis('off')
plt.show()
