from streamlit.testing.v1 import AppTest


def main():
    app = AppTest.from_file("streamlit_app.py")
    app.run(timeout=360)
    if app.exception:
        for exception in app.exception:
            print(exception.value)
        raise SystemExit(f"Streamlit smoke test failed with {len(app.exception)} exception(s).")
    print("Streamlit smoke test passed.")


if __name__ == "__main__":
    main()
