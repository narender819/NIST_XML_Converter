import numpy as np

class CorrelationEngine:
    """
    Production-grade Thermodynamic Correlation Engine

    Supports:
    - Equation 1
    - Equation 18
    - First derivative
    """

    def __init__(self):
        pass

    # =========================================================
    # Public Interface
    # =========================================================
    def evaluate(self, eq_no, T, coeffs, derivative=False):
        """
        Evaluate correlation or its first derivative.

        Parameters
        ----------
        eq_no : int
            Equation number
        T : float or array-like
            Temperature in K
        coeffs : list or array
            Coefficients [C1, C2, ..., Cn]
        derivative : bool
            If True, returns dH/dT

        Returns
        -------
        float or np.ndarray
        """
        T = self._prepare_temperature(T)
        coeffs = self._prepare_coefficients(coeffs)

        if eq_no == 1:
            return self._eq1(T, coeffs, derivative)
        elif eq_no == 18:
            return self._eq18(T, coeffs, derivative)
        else:
            raise NotImplementedError(f"Equation {eq_no} not implemented")

    # =========================================================
    # Equation 1 (Polynomial)
    # H = C1 + C2*T + C3*T^2 + C4*T^3 + ...
    # =========================================================
    def _eq1(self, T, coeffs, derivative=False):
        if not derivative:
            return np.polynomial.polynomial.polyval(T, coeffs)
        else:
            deriv_coeffs = [i * coeffs[i] for i in range(1, len(coeffs))]
            return np.polynomial.polynomial.polyval(T, deriv_coeffs)

    # =========================================================
    # Equation 18
    # Prop = C1 + (C2 * T^C3) / (1 + C4/T + C5/T^2)
    # =========================================================
    def _eq18(self, T, coeffs, derivative=False):
        T = np.asarray(T, dtype=float)

        # Ensure at least 5 coefficients, pad with zeros if needed
        C = list(coeffs) + [0.0] * (5 - len(coeffs))
        C1, C2, C3, C4, C5 = C[:5]

        # Common terms
        T_pow_C3 = T ** C3
        inv_T = 1.0 / T
        inv_T2 = inv_T ** 2

        denom = 1.0 + C4 * inv_T + C5 * inv_T2

        if not derivative:
            # H(T)
            return C1 + (C2 * T_pow_C3) / denom
        else:
            # dH/dT
            # f(T) = C2*T^C3 / denom
            # f' = [C2*C3*T^(C3-1)*denom - C2*T^C3*denom'] / denom^2
            denom_prime = -C4 * inv_T2 - 2.0 * C5 * inv_T**3
            num = C2 * C3 * T**(C3 - 1) * denom - C2 * T_pow_C3 * denom_prime
            dH = num / (denom ** 2)
            return dH

    # =========================================================
    # Helpers
    # =========================================================
    def _prepare_temperature(self, T):
        if isinstance(T, (list, tuple)):
            return np.array(T, dtype=float)
        elif isinstance(T, np.ndarray):
            return T.astype(float)
        else:
            return float(T)

    def _prepare_coefficients(self, coeffs):
        coeffs = list(coeffs)
        clean = []
        for c in coeffs:
            if c is None:
                clean.append(0.0)
            else:
                clean.append(float(c))
        return clean
