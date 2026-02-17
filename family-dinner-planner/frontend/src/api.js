const API_BASE_URL = import.meta.env.VITE_API_URL || "http://localhost:5000/api";

async function request(path, options = {}) {
  const { token, ...rest } = options;
  const response = await fetch(`${API_BASE_URL}${path}`, {
    ...rest,
    headers: {
      "Content-Type": "application/json",
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...(rest.headers || {})
    }
  });

  const payload = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(payload.message || "Request failed");
  }

  return payload;
}

export const api = {
  register(data) {
    return request("/auth/register", {
      method: "POST",
      body: JSON.stringify(data)
    });
  },

  login(data) {
    return request("/auth/login", {
      method: "POST",
      body: JSON.stringify(data)
    });
  },

  getDinners(token) {
    return request("/dinners", { token });
  },

  addDinner(token, data) {
    return request("/dinners", {
      method: "POST",
      token,
      body: JSON.stringify(data)
    });
  },

  submitWeeklyChoices(token, data) {
    return request("/submissions", {
      method: "POST",
      token,
      body: JSON.stringify(data)
    });
  },

  getSubmissions(token) {
    return request("/submissions", { token });
  },

  getMySubmissions(token) {
    return request("/submissions/mine", { token });
  }
};
