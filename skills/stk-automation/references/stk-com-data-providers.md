# STK COM Data Provider Notes

Use these notes when implementing or reviewing STK automation scripts.

## Object Creation Enums Observed in STK 11

- `1`: Aircraft
- `4`: Chain
- `8`: Facility
- `17`: Receiver, as a child object
- `24`: Transmitter, as a child object

Aircraft external ephemeris route type observed in STK 11:

- `aircraft.SetRouteType(6)`
- `aircraft.Route.Filename = <path-to-e-file>`
- `aircraft.Route.Propagate()`

Great Arc route type observed in STK 11:

- `aircraft.SetRouteType(9)`

## Reliable Export Providers

Aircraft state:

- `aircraft.DataProviders.Item("Cartesian Position").Group.Item("Fixed")`
- `aircraft.DataProviders.Item("Cartesian Velocity").Group.Item("Fixed")`
- `aircraft.DataProviders.Item("Cartesian Acceleration").Group.Item("Fixed")`

Link data:

- `access.DataProviders.Item("Link Information")`
- Doppler field: `Freq. Doppler Shift`

Doppler rate:

1. Create a VGT scalar from the link-information data element:
   `access.Vgt.CalcScalars.Factory.CreateCalcScalarDataElement(name, desc, "Link Information", "Freq. Doppler Shift")`
2. Export:
   `access.DataProviders.Item("Scalar Calculations").Group.Item(name)`
3. Use the returned `Scalar Rate` column.

## Quality Lessons

Multi-waypoint paths can look acceptable in STK while producing bad data at segment boundaries. If Doppler-rate is a deliverable, prefer single-segment paths, Aviator, or external ephemeris generated from a smooth curve and parameterized by arc length.

