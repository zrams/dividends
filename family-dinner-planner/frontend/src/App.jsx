import { useEffect, useMemo, useState } from "react";
import { api } from "./api";

function getTodayInputValue() {
  return new Date().toISOString().slice(0, 10);
}

function formatDate(value) {
  return new Date(value).toLocaleDateString();
}

export default function App() {
  const [authMode, setAuthMode] = useState("login");
  const [authForm, setAuthForm] = useState({
    name: "",
    email: "",
    password: "",
    role: "member"
  });
  const [session, setSession] = useState(() => {
    const raw = localStorage.getItem("fdp_session");
    return raw ? JSON.parse(raw) : null;
  });
  const [dinners, setDinners] = useState([]);
  const [submissions, setSubmissions] = useState([]);
  const [mySubmissions, setMySubmissions] = useState([]);
  const [dinnerForm, setDinnerForm] = useState({ name: "", description: "" });
  const [selectedChoices, setSelectedChoices] = useState([]);
  const [weekStartDate, setWeekStartDate] = useState(getTodayInputValue());
  const [statusMessage, setStatusMessage] = useState("");
  const [errorMessage, setErrorMessage] = useState("");
  const [isLoading, setIsLoading] = useState(false);

  const isAdmin = session?.user?.role === "admin";
  const isMember = session?.user?.role === "member";

  const sortedDinners = useMemo(
    () => [...dinners].sort((a, b) => a.name.localeCompare(b.name)),
    [dinners]
  );

  useEffect(() => {
    if (!session) {
      return;
    }

    async function loadData() {
      try {
        setIsLoading(true);
        setErrorMessage("");
        const dinnerList = await api.getDinners(session.token);
        setDinners(dinnerList);

        if (session.user.role === "admin") {
          const submissionList = await api.getSubmissions(session.token);
          setSubmissions(submissionList);
        } else {
          const ownSubmissions = await api.getMySubmissions(session.token);
          setMySubmissions(ownSubmissions);
        }
      } catch (error) {
        setErrorMessage(error.message);
      } finally {
        setIsLoading(false);
      }
    }

    loadData();
  }, [session]);

  function handleAuthFormChange(event) {
    const { name, value } = event.target;
    setAuthForm((prev) => ({ ...prev, [name]: value }));
  }

  async function handleAuthSubmit(event) {
    event.preventDefault();
    setErrorMessage("");
    setStatusMessage("");

    try {
      setIsLoading(true);
      const payload =
        authMode === "register"
          ? await api.register(authForm)
          : await api.login({ email: authForm.email, password: authForm.password });

      setSession(payload);
      localStorage.setItem("fdp_session", JSON.stringify(payload));
      setStatusMessage(`Welcome, ${payload.user.name}!`);
      setSelectedChoices([]);
      setAuthForm((prev) => ({ ...prev, password: "" }));
    } catch (error) {
      setErrorMessage(error.message);
    } finally {
      setIsLoading(false);
    }
  }

  async function handleAddDinner(event) {
    event.preventDefault();
    setErrorMessage("");
    setStatusMessage("");

    try {
      setIsLoading(true);
      const createdDinner = await api.addDinner(session.token, dinnerForm);
      setDinners((prev) => [createdDinner, ...prev]);
      setDinnerForm({ name: "", description: "" });
      setStatusMessage("Dinner idea added.");
    } catch (error) {
      setErrorMessage(error.message);
    } finally {
      setIsLoading(false);
    }
  }

  function handleToggleChoice(dinnerId) {
    setSelectedChoices((prev) => {
      if (prev.includes(dinnerId)) {
        return prev.filter((id) => id !== dinnerId);
      }

      if (prev.length === 2) {
        setErrorMessage("You can only select 2 choices.");
        return prev;
      }

      setErrorMessage("");
      return [...prev, dinnerId];
    });
  }

  async function handleSubmitChoices(event) {
    event.preventDefault();
    setErrorMessage("");
    setStatusMessage("");

    try {
      setIsLoading(true);
      await api.submitWeeklyChoices(session.token, {
        weekStartDate,
        choices: selectedChoices
      });
      const ownSubmissions = await api.getMySubmissions(session.token);
      setMySubmissions(ownSubmissions);
      setSelectedChoices([]);
      setStatusMessage("Weekly choices submitted.");
    } catch (error) {
      setErrorMessage(error.message);
    } finally {
      setIsLoading(false);
    }
  }

  function handleLogout() {
    setSession(null);
    localStorage.removeItem("fdp_session");
    setDinners([]);
    setSubmissions([]);
    setMySubmissions([]);
    setStatusMessage("Logged out.");
    setErrorMessage("");
  }

  if (!session) {
    return (
      <main className="container">
        <h1>Family Dinner Planner</h1>
        <p>Plan weekly dinners together with admin/member roles.</p>

        <section className="card">
          <div className="auth-toggle">
            <button
              type="button"
              className={authMode === "login" ? "active" : ""}
              onClick={() => setAuthMode("login")}
            >
              Login
            </button>
            <button
              type="button"
              className={authMode === "register" ? "active" : ""}
              onClick={() => setAuthMode("register")}
            >
              Register
            </button>
          </div>

          <form onSubmit={handleAuthSubmit} className="form">
            {authMode === "register" && (
              <>
                <label>
                  Name
                  <input
                    name="name"
                    value={authForm.name}
                    onChange={handleAuthFormChange}
                    required
                  />
                </label>
                <label>
                  Role
                  <select name="role" value={authForm.role} onChange={handleAuthFormChange}>
                    <option value="member">member</option>
                    <option value="admin">admin (mom)</option>
                  </select>
                </label>
              </>
            )}
            <label>
              Email
              <input
                type="email"
                name="email"
                value={authForm.email}
                onChange={handleAuthFormChange}
                required
              />
            </label>
            <label>
              Password
              <input
                type="password"
                name="password"
                value={authForm.password}
                onChange={handleAuthFormChange}
                required
                minLength={6}
              />
            </label>

            <button type="submit" disabled={isLoading}>
              {isLoading ? "Please wait..." : authMode === "register" ? "Create account" : "Sign in"}
            </button>
          </form>
        </section>

        {errorMessage && <p className="error">{errorMessage}</p>}
        {statusMessage && <p className="success">{statusMessage}</p>}
      </main>
    );
  }

  return (
    <main className="container">
      <header className="header">
        <div>
          <h1>Family Dinner Planner</h1>
          <p>
            Signed in as <strong>{session.user.name}</strong> ({session.user.role})
          </p>
        </div>
        <button type="button" onClick={handleLogout}>
          Logout
        </button>
      </header>

      {errorMessage && <p className="error">{errorMessage}</p>}
      {statusMessage && <p className="success">{statusMessage}</p>}

      <section className="card">
        <h2>Available Dinner Ideas ({dinners.length})</h2>
        {sortedDinners.length === 0 ? (
          <p>No dinner ideas yet. {isAdmin ? "Add your first one below." : "Ask admin to add ideas."}</p>
        ) : (
          <ul className="list">
            {sortedDinners.map((dinner) => (
              <li key={dinner._id}>
                <strong>{dinner.name}</strong>
                {dinner.description ? ` - ${dinner.description}` : ""}
              </li>
            ))}
          </ul>
        )}
      </section>

      {isAdmin && (
        <>
          <section className="card">
            <h2>Add Dinner Idea</h2>
            <form onSubmit={handleAddDinner} className="form">
              <label>
                Name
                <input
                  value={dinnerForm.name}
                  onChange={(event) =>
                    setDinnerForm((prev) => ({ ...prev, name: event.target.value }))
                  }
                  required
                />
              </label>
              <label>
                Description
                <textarea
                  value={dinnerForm.description}
                  onChange={(event) =>
                    setDinnerForm((prev) => ({ ...prev, description: event.target.value }))
                  }
                />
              </label>
              <button type="submit" disabled={isLoading}>
                Add dinner
              </button>
            </form>
            <p className="muted">
              Tip: Add 50-100 dinner options so family members can vote from a strong list.
            </p>
          </section>

          <section className="card">
            <h2>Weekly Submissions</h2>
            {submissions.length === 0 ? (
              <p>No member submissions yet.</p>
            ) : (
              <ul className="list">
                {submissions.map((submission) => (
                  <li key={submission._id}>
                    <div>
                      <strong>{submission.userId?.name || "Unknown user"}</strong> (
                      {submission.userId?.email || "no email"}) - week of{" "}
                      {formatDate(submission.weekStartDate)}
                    </div>
                    <ul>
                      {submission.choices.map((choice) => (
                        <li key={choice._id}>{choice.name}</li>
                      ))}
                    </ul>
                  </li>
                ))}
              </ul>
            )}
          </section>
        </>
      )}

      {isMember && (
        <>
          <section className="card">
            <h2>Submit This Week's Top 1-2 Choices</h2>
            <form onSubmit={handleSubmitChoices} className="form">
              <label>
                Week date
                <input
                  type="date"
                  value={weekStartDate}
                  onChange={(event) => setWeekStartDate(event.target.value)}
                  required
                />
              </label>
              <div>
                <p>Select up to 2 dinners:</p>
                <ul className="choice-list">
                  {sortedDinners.map((dinner) => (
                    <li key={dinner._id}>
                      <label>
                        <input
                          type="checkbox"
                          checked={selectedChoices.includes(dinner._id)}
                          onChange={() => handleToggleChoice(dinner._id)}
                        />
                        {dinner.name}
                      </label>
                    </li>
                  ))}
                </ul>
              </div>
              <button
                type="submit"
                disabled={selectedChoices.length < 1 || selectedChoices.length > 2 || isLoading}
              >
                Submit choices
              </button>
            </form>
          </section>

          <section className="card">
            <h2>My Past Submissions</h2>
            {mySubmissions.length === 0 ? (
              <p>No submissions yet.</p>
            ) : (
              <ul className="list">
                {mySubmissions.map((submission) => (
                  <li key={submission._id}>
                    Week of {formatDate(submission.weekStartDate)}:
                    <ul>
                      {submission.choices.map((choice) => (
                        <li key={choice._id}>{choice.name}</li>
                      ))}
                    </ul>
                  </li>
                ))}
              </ul>
            )}
          </section>
        </>
      )}
    </main>
  );
}
