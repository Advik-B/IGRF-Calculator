import os
import math
import pandas as pd
from datetime import date, datetime


class GeoMag:
    def __init__(self, coeffs_filename=None, desired_year=None):
        """
        Initialize the GeoMag object with updated IGRF-14 coefficients from an Excel file.

        Parameters:
            coeffs_filename (str): Path to the Excel file containing updated coefficients.
            desired_year (float): The epoch (decimal year) for which to interpolate the coefficients.
                                   If None, defaults to the latest epoch available (e.g. 2025).
        """
        if not coeffs_filename:
            coeffs_filename = os.path.join(os.path.dirname(__file__), 'coeff.xlsx')

        # Read file
        if coeffs_filename.endswith(('.xls', '.xlsx')):
            coeffs_df = pd.read_excel(coeffs_filename, sheet_name='igrf14coeffs')
        else:
            coeffs_df = pd.read_csv(coeffs_filename)

        # Debug: print column names so you can verify them.
        print("Excel columns:", coeffs_df.columns.tolist())

        # Expecting columns: ["g/h", "n", "m", "1900", "1905", ..., "2025", "2025-30"]
        # Extract epoch columns (from index 3 to second-to-last).
        all_cols = list(coeffs_df.columns)
        epoch_col_names = all_cols[3:-1]
        # Convert these column names to floats.
        available_epochs = []
        for col in epoch_col_names:
            try:
                available_epochs.append(float(col))
            except ValueError:
                pass
        available_epochs.sort()
        print("Available epochs:", available_epochs)

        # Set desired_year. If not provided, default to the last available epoch.
        if desired_year is None:
            desired_year = available_epochs[-1]
        self.desired_year = desired_year

        # Now, build coefficient dictionaries for g and h.
        # For each row in the Excel file, perform linear interpolation of the coefficient values.
        g_dict = {}
        h_dict = {}

        for _, row in coeffs_df.iterrows():
            typ = str(row["g/h"]).strip().lower()  # 'g' or 'h'
            n = int(row["n"])
            m = int(row["m"])
            # Build list of (epoch, value) for this row:
            values = []
            for col in epoch_col_names:
                values.append(float(row[col]))
            # Interpolate coefficient for the desired_year.
            coef = self.interpolate_coef(available_epochs, values, desired_year)
            # Also, you might want to extract secular variation from the last column if needed:
            sec_var = float(row[all_cols[-1]])

            if typ == 'g':
                g_dict[(n, m)] = (coef, sec_var)
            elif typ == 'h':
                h_dict[(n, m)] = (coef, sec_var)

        # Store the coefficients in internal arrays in the same way your existing code does.
        self.maxord = self.maxdeg = 12  # or adjust if needed
        z = [0.0] * 14
        self.tc = [z[:] for _ in range(14)]
        self.sp = z[:]
        self.cp = z[:]
        self.cp[0] = 1.0
        self.pp = z[:13]
        self.pp[0] = 1.0
        self.p = [z[:] for _ in range(14)]
        self.p[0][0] = 1.0
        self.dp = [z[:] for _ in range(13)]
        self.a = 6378.137
        self.b = 6356.7523142
        self.re = 6371.2
        self.a2 = self.a * self.a
        self.b2 = self.b * self.b
        self.c2 = self.a2 - self.b2
        self.a4 = self.a2 * self.a2
        self.b4 = self.b2 * self.b2
        self.c4 = self.a4 - self.b4

        self.c = [z[:] for _ in range(14)]
        self.cd = [z[:] for _ in range(14)]

        # Populate coefficient arrays.
        # For each coefficient (n,m), if it is a g coefficient, place it in self.c[m][n],
        # and if m != 0, place the corresponding h coefficient in self.c[n][m-1].
        for key, (val, sec) in g_dict.items():
            n, m = key
            if m <= n:
                self.c[m][n] = val
                self.cd[m][n] = sec
        for key, (val, sec) in h_dict.items():
            n, m = key
            if m <= n and m != 0:
                self.c[n][m - 1] = val
                self.cd[n][m - 1] = sec

        # Convert Schmidt normalized Gauss coefficients to unnormalized.
        self.snorm = [z[:] for _ in range(13)]
        self.snorm[0][0] = 1.0
        self.k = [z[:] for _ in range(13)]
        self.k[1][1] = 0.0
        self.fn = [0.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0, 13.0]
        self.fm = [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.0, 12.0]
        for n in range(1, self.maxord + 1):
            self.snorm[0][n] = self.snorm[0][n - 1] * (2.0 * n - 1) / n
            j = 2.0
            m = 0
            D1 = 1
            D2 = (n - m + 1)
            while D2 > 0:
                self.k[m][n] = (((n - 1) ** 2) - (m * m)) / ((2.0 * n - 1) * (2.0 * n - 3.0))
                if m > 0:
                    flnmj = ((n - m + 1.0) * j) / (n + m)
                    self.snorm[m][n] = self.snorm[m - 1][n] * math.sqrt(flnmj)
                    j = 1.0
                    self.c[n][m - 1] = self.snorm[m][n] * self.c[n][m - 1]
                    self.cd[n][m - 1] = self.snorm[m][n] * self.cd[n][m - 1]
                self.c[m][n] = self.snorm[m][n] * self.c[m][n]
                self.cd[m][n] = self.snorm[m][n] * self.cd[m][n]
                D2 = D2 - 1
                m = m + D1

        self.epoch = desired_year

    def interpolate_coef(self, epochs, values, target):
        """
        Linearly interpolate (or extrapolate) a coefficient value.

        Parameters:
            epochs (list of float): Sorted list of epochs.
            values (list of float): Corresponding coefficient values.
            target (float): Desired epoch (decimal year).

        Returns:
            float: Interpolated coefficient.
        """
        # If target is outside the available range, extrapolate using the first two or last two points.
        if target <= epochs[0]:
            x0, x1 = epochs[0], epochs[1]
            y0, y1 = values[0], values[1]
        elif target >= epochs[-1]:
            x0, x1 = epochs[-2], epochs[-1]
            y0, y1 = values[-2], values[-1]
        else:
            # Find indices such that x0 <= target <= x1.
            for i in range(len(epochs) - 1):
                if epochs[i] <= target <= epochs[i + 1]:
                    x0, x1 = epochs[i], epochs[i + 1]
                    y0, y1 = values[i], values[i + 1]
                    break
        # Linear interpolation:
        if (x1 - x0) == 0:
            return y0
        return y0 + (y1 - y0) * (target - x0) / (x1 - x0)

    def GeoMag(self, dlat, dlon, h=0,
               time=date.today()):  # latitude (decimal degrees), longitude (decimal degrees), altitude (feet), date
        # time = date('Y') + date('z')/365
        time = time.year + ((time - date(time.year, 1, 1)).days / 365.0)
        alt = h / 3280.8399

        otime = oalt = olat = olon = -1000.0

        dt = time - self.epoch
        glat = dlat
        glon = dlon
        rlat = math.radians(glat)
        rlon = math.radians(glon)
        srlon = math.sin(rlon)
        srlat = math.sin(rlat)
        crlon = math.cos(rlon)
        crlat = math.cos(rlat)
        srlat2 = srlat * srlat
        crlat2 = crlat * crlat
        self.sp[1] = srlon
        self.cp[1] = crlon

        # /* CONVERT FROM GEODETIC COORDS. TO SPHERICAL COORDS. */
        if (alt != oalt or glat != olat):
            q = math.sqrt(self.a2 - self.c2 * srlat2)
            q1 = alt * q
            q2 = ((q1 + self.a2) / (q1 + self.b2)) * ((q1 + self.a2) / (q1 + self.b2))
            ct = srlat / math.sqrt(q2 * crlat2 + srlat2)
            st = math.sqrt(1.0 - (ct * ct))
            r2 = (alt * alt) + 2.0 * q1 + (self.a4 - self.c4 * srlat2) / (q * q)
            r = math.sqrt(r2)
            d = math.sqrt(self.a2 * crlat2 + self.b2 * srlat2)
            ca = (alt + d) / r
            sa = self.c2 * crlat * srlat / (r * d)

        if (glon != olon):
            for m in range(2, self.maxord + 1):
                self.sp[m] = self.sp[1] * self.cp[m - 1] + self.cp[1] * self.sp[m - 1]
                self.cp[m] = self.cp[1] * self.cp[m - 1] - self.sp[1] * self.sp[m - 1]

        aor = self.re / r
        ar = aor * aor
        br = bt = bp = bpp = 0.0
        for n in range(1, self.maxord + 1):
            ar = ar * aor

            # for (m=0,D3=1,D4=(n+m+D3)/D3;D4>0;D4--,m+=D3):
            m = 0
            D3 = 1
            # D4=(n+m+D3)/D3
            D4 = (n + m + 1)
            while D4 > 0:

                # /*
                # COMPUTE UNNORMALIZED ASSOCIATED LEGENDRE POLYNOMIALS
                # AND DERIVATIVES VIA RECURSION RELATIONS
                # */
                if (alt != oalt or glat != olat):
                    if (n == m):
                        self.p[m][n] = st * self.p[m - 1][n - 1]
                        self.dp[m][n] = st * self.dp[m - 1][n - 1] + ct * self.p[m - 1][n - 1]

                    elif (n == 1 and m == 0):
                        self.p[m][n] = ct * self.p[m][n - 1]
                        self.dp[m][n] = ct * self.dp[m][n - 1] - st * self.p[m][n - 1]

                    elif (n > 1 and n != m):
                        if (m > n - 2):
                            self.p[m][n - 2] = 0
                        if (m > n - 2):
                            self.dp[m][n - 2] = 0.0
                        self.p[m][n] = ct * self.p[m][n - 1] - self.k[m][n] * self.p[m][n - 2]
                        self.dp[m][n] = ct * self.dp[m][n - 1] - st * self.p[m][n - 1] - self.k[m][n] * self.dp[m][
                            n - 2]

                # /*
                # TIME ADJUST THE GAUSS COEFFICIENTS
                # */
                if (time != otime):
                    self.tc[m][n] = self.c[m][n] + dt * self.cd[m][n]
                    if (m != 0):
                        self.tc[n][m - 1] = self.c[n][m - 1] + dt * self.cd[n][m - 1]

                # /*
                # ACCUMULATE TERMS OF THE SPHERICAL HARMONIC EXPANSIONS
                # */
                par = ar * self.p[m][n]

                if (m == 0):
                    temp1 = self.tc[m][n] * self.cp[m]
                    temp2 = self.tc[m][n] * self.sp[m]
                else:
                    temp1 = self.tc[m][n] * self.cp[m] + self.tc[n][m - 1] * self.sp[m]
                    temp2 = self.tc[m][n] * self.sp[m] - self.tc[n][m - 1] * self.cp[m]

                bt = bt - ar * temp1 * self.dp[m][n]
                bp = bp + (self.fm[m] * temp2 * par)
                br = br + (self.fn[n] * temp1 * par)
                # /*
                # SPECIAL CASE:  NORTH/SOUTH GEOGRAPHIC POLES
                # */
                if (st == 0.0 and m == 1):
                    if (n == 1):
                        self.pp[n] = self.pp[n - 1]
                    else:
                        self.pp[n] = ct * self.pp[n - 1] - self.k[m][n] * self.pp[n - 2]
                    parp = ar * self.pp[n]
                    bpp = bpp + (self.fm[m] * temp2 * parp)

                D4 = D4 - 1
                m = m + 1

        if (st == 0.0):
            bp = bpp
        else:
            bp = bp / st
        # /*
        # ROTATE MAGNETIC VECTOR COMPONENTS FROM SPHERICAL TO
        # GEODETIC COORDINATES
        # */
        bx = -bt * ca - br * sa
        by = bp
        bz = bt * sa - br * ca
        # /*
        # COMPUTE DECLINATION (DEC), INCLINATION (DIP) AND
        # TOTAL INTENSITY (TI)
        # */
        bh = math.sqrt((bx * bx) + (by * by))
        ti = math.sqrt((bh * bh) + (bz * bz))
        dec = math.degrees(math.atan2(by, bx))
        dip = math.degrees(math.atan2(bz, bh))
        # /*
        # COMPUTE MAGNETIC GRID VARIATION IF THE CURRENT
        # GEODETIC POSITION IS IN THE ARCTIC OR ANTARCTIC
        # (I.E. GLAT > +55 DEGREES OR GLAT < -55 DEGREES)

        # OTHERWISE, SET MAGNETIC GRID VARIATION TO -999.0
        # */
        gv = -999.0
        if (math.fabs(glat) >= 55.):
            if (glat > 0.0 and glon >= 0.0):
                gv = dec - glon
            if (glat > 0.0 and glon < 0.0):
                gv = dec + math.fabs(glon);
            if (glat < 0.0 and glon >= 0.0):
                gv = dec + glon
            if (glat < 0.0 and glon < 0.0):
                gv = dec - math.fabs(glon)
            if (gv > +180.0):
                gv = gv - 360.0
            if (gv < -180.0):
                gv = gv + 360.0

        otime = time
        oalt = alt
        olat = glat
        olon = glon

        class RetObj:
            pass

        retobj = RetObj()
        retobj.dec = dec
        retobj.dip = dip
        retobj.ti = ti
        retobj.bh = bh
        retobj.bx = bx
        retobj.by = by
        retobj.bz = bz
        retobj.lat = dlat
        retobj.lon = dlon
        retobj.alt = h
        retobj.time = time

        return retobj


