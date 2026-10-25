# Sourced by export-span-traces.sh during tests and by operators.
# A real script from the other team would refresh AWS credentials here.
# This stand-in only records that it was sourced and clears the fake denial.
if [[ -n "${FAKE_AWS_DIR:-}" ]]; then
  printf 'sourced\n' >> "${FAKE_AWS_DIR}/switchback.log"
fi
export FAKE_AWS_MODE=ok
