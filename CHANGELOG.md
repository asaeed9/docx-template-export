# Changelog

All notable changes to this project will be documented in this file.

## v0.2.0

### Major Enhancements

- Safe nested Word table rendering with atomic commit and deterministic fallback
- Deterministic table formatting using Word styles with XML border fallback
- Multi-run–safe scalar placeholder replacement in headers, footers, and textboxes

### Safety & Guardrails

- Atomic block placeholder expansion (freeze ? validate ? mutate)
- Strict BODY-only block expansion rules
- Guaranteed structural validity of table cells after replacement

### Internal Improvements

- Removal of unused legacy helpers and added invariant documentation
- Improved code organization with separated table formatting responsibilities
- Simplified cell detection logic for better maintainability

### Backward Compatibility

- No breaking API changes; fully backward compatible
- All existing functionality preserved with improved safety guarantees
