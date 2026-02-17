// Capstone Template Structure
// Defines the sections, subsections, and sub-subsections for the Calculations Sheet

export const capstoneTemplate = {
  id: 'capstone',
  name: 'Capstone Template',
  columns: [
    'Estimate',
    'Particulars',
    'Takeoff',
    'Unit',
    'QTY',
    'Length',
    'Width',
    'Height',
    'FT',
    'SQ FT',
    'LBS',
    'CY',
    'QTY'
  ],
  structure: [
    {
      section: 'Demolition',
      subsections: [
        { name: 'Demo slab on grade', items: [] },
        { name: 'Demo Ramp on grade', items: [] },
        { name: 'Demo strip footing', items: [] },
        { name: 'Demo foundation wall', items: [] },
        { name: 'Demo retaining wall', items: [] },
        { name: 'Demo isolated footing', items: [] }
      ]
    },
    {
      section: 'Excavation',
      subsections: [
        { name: 'Excavation', items: [] },
        { name: 'Backfill', items: [] },
        { name: 'Mud slab', items: [] }
      ]
    },
    {
      section: 'Rock Excavation',
      subsections: [
        { name: 'Excavation', items: [] },
        { name: 'Line drill', items: [] }
      ]
    },
    {
      section: 'SOE',
      subsections: [
        { name: 'Drilled soldier pile', items: [] },
        { name: 'Primary secant piles', items: [] },
        { name: 'Secondary secant piles', items: [] },
        { name: 'Tangent piles', items: [] },
        { name: 'Sheet pile', items: [] },
        { name: 'Timber lagging', items: [] },
        { name: 'Backpacking', items: [] },
        { name: 'Timber sheeting', items: [] },
        { name: 'Timber soldier piles', items: [] },
        { name: 'Timber planks', items: [] },
        { name: 'Timber waler', items: [] },
        { name: 'Timber raker', items: [] },
        { name: 'Timber brace', items: [] },
        { name: 'Timber post', items: [] },
        { name: 'Vertical timber sheets', items: [] },
        { name: 'Horizontal timber sheets', items: [] },
        { name: 'Timber stringer', items: [] },
        { name: 'Waler', items: [] },
        { name: 'Raker', items: [] },
        { name: 'Upper Raker', items: [] },
        { name: 'Lower Raker', items: [] },
        { name: 'Stand off', items: [] },
        { name: 'Kicker', items: [] },
        { name: 'Channel', items: [] },
        { name: 'Roll chock', items: [] },
        { name: 'Stud beam', items: [] },
        { name: 'Inner corner brace', items: [] },
        { name: 'Knee brace', items: [] },
        { name: 'Supporting angle', items: [] },
        { name: 'Parging', items: [] },
        { name: 'Heel blocks', items: [] },
        { name: 'Underpinning', items: [] },
        { name: 'Rock anchors', items: [] },
        { name: 'Rock bolts', items: [] },
        { name: 'Anchor', items: [] },
        { name: 'Tie back', items: [] },
        { name: 'Concrete soil retention piers', items: [] },
        { name: 'Guide wall', items: [] },
        { name: 'Dowel bar', items: [] },
        { name: 'Rock pins', items: [] },
        { name: 'Shotcrete', items: [] },
        { name: 'Permission grouting', items: [] },
        { name: 'Buttons', items: [] },
        { name: 'Rock stabilization', items: [] },
        { name: 'Form board', items: [] },
        { name: 'Drilled hole grout', items: [] }
      ]
    },
    {
      section: 'Foundation',
      subsections: [
        { name: 'Piles', items: [] },
        { name: 'Drilled foundation pile', items: [] },
        { name: 'Helical foundation pile', items: [] },
        { name: 'Driven foundation pile', items: [] },
        { name: 'Stelcor drilled displacement pile', items: [] },
        { name: 'CFA pile', items: [] },
        { name: 'Pile caps', items: [] },
        { name: 'Strip Footings', items: [] },
        { name: 'Isolated Footings', items: [] },
        { name: 'Pilaster', items: [] },
        { name: 'Grade beams', items: [] },
        { name: 'Tie beam', items: [] },
        { name: 'Strap beams', items: [] },
        { name: 'Thickened slab', items: [] },
        { name: 'Buttresses', items: [] },
        { name: 'Pier', items: [] },
        { name: 'Corbel', items: [] },
        { name: 'Linear Wall', items: [] },
        { name: 'Foundation Wall', items: [] },
        { name: 'Retaining walls', items: [] },
        { name: 'Barrier wall', items: [] },
        { name: 'Stem wall', items: [] },
        { name: 'Elevator Pit', items: [] },
        { name: 'Service elevator pit', items: [] },
        { name: 'Detention tank', items: [] },
        { name: 'Duplex sewage ejector pit', items: [] },
        { name: 'Deep sewage ejector pit', items: [] },
        { name: 'Sump pump pit', items: [] },
        { name: 'Grease trap', items: [] },
        { name: 'House trap', items: [] },
        { name: 'Mat slab', items: [] },
        { name: 'Mud Slab', items: [] },
        { name: 'SOG', items: [] },
        { name: 'Stairs on grade Stairs', items: [] },
        { name: 'Electric conduit', items: [] }
      ]
    },
    {
      section: 'Waterproofing',
      subsections: [
        { name: 'Exterior side', items: [] },
        { name: 'Negative side', items: [] },
        { name: 'Horizontal', items: [] }
      ]
    },
    {
      section: 'Trenching',
      subsections: []
    },
    {
      section: 'Superstructure',
      subsections: [
        { name: 'CIP Slabs', items: [] },
        { name: 'Balcony slab', items: [] },
        { name: 'Terrace slab', items: [] },
        { name: 'Patch slab', items: [] },
        { name: 'Slab steps', items: [] },
        { name: 'LW concrete fill', items: [] },
        { name: 'Slab on metal deck', items: [] },
        { name: 'Topping slab', items: [] },
        { name: 'Thermal break', items: [] },
        { name: 'Raised slab', items: [] },
        { name: 'Built-up slab', items: [] },
        { name: 'Builtup ramps', items: [] },
        { name: 'Built-up stair', items: [] },
        { name: 'Concrete hanger', items: [] },
        { name: 'Shear Walls', items: [] },
        { name: 'Parapet walls', items: [] },
        { name: 'Columns', items: [] },
        { name: 'Concrete post', items: [] },
        { name: 'Concrete encasement', items: [] },
        { name: 'Drop panel', items: [] },
        { name: 'Beams', items: [] },
        { name: 'CIP Stairs', items: [] },
        { name: 'Stairs – Infilled tads', items: [] },
        { name: 'Curbs', items: [] },
        { name: 'Concrete pad', items: [] },
        { name: 'Non-shrink grout', items: [] },
        { name: 'Repair scope', items: [] }
      ]
    },
    {
      section: 'B.P.P. Alternate #2 scope',
      subsections: [
        { name: 'Gravel', items: [] },
        { name: 'Concrete sidewalk', items: [] },
        { name: 'Concrete driveway', items: [] },
        { name: 'Concrete curb', items: [] },
        { name: 'Concrete flush curb', items: [] },
        { name: 'Expansion joint', items: [] },
        { name: 'Conc road base', items: [] },
        { name: 'Full depth asphalt pavement', items: [] }
      ]
    },
    {
      section: 'Civil / Sitework',
      subsections: [
        {
          name: 'Demo',
          subSubsections: [
            { name: 'Demo asphalt', items: [] },
            { name: 'Demo curb', items: [] },
            { name: 'Demo fence', items: [] },
            { name: 'Demo wall', items: [] },
            { name: 'Demo pipe', items: [] },
            { name: 'Demo rail', items: [] },
            { name: 'Demo sign', items: [] },
            { name: 'Demo manhole', items: [] },
            { name: 'Demo fire hydrant', items: [] },
            { name: 'Demo utility pole', items: [] },
            { name: 'Demo valve', items: [] },
            { name: 'Demo inlet', items: [] }
          ]
        },
        { name: 'Excavation', items: [] },
        { name: 'Gravel', items: [] },
        { name: 'Concrete Pavement', items: [] },
        { name: 'Asphalt', items: [] },
        { name: 'Pads', items: [] },
        { name: 'Soil Erosion', items: [] },
        { name: 'Fence', items: [] },
        { name: 'Concrete filled steel pipe bollard', items: [] },
        {
          name: 'Site',
          subSubsections: [
            { name: 'Hydrant', items: [] },
            { name: 'Wheel stop', items: [] },
            { name: 'Drain', items: [] },
            { name: 'Protection', items: [] },
            { name: 'Signages', items: [] },
            { name: 'Main line', items: [] }
          ]
        },
        {
          name: 'Ele',
          subSubsections: [
            { name: 'Excavation', items: [] },
            { name: 'Backfill', items: [] },
            { name: 'Gravel', items: [] }
          ]
        },
        {
          name: 'Gas',
          subSubsections: [
            { name: 'Excavation', items: [] },
            { name: 'Backfill', items: [] },
            { name: 'Gravel', items: [] }
          ]
        },
        {
          name: 'Water',
          subSubsections: [
            { name: 'Excavation', items: [] },
            { name: 'Backfill', items: [] },
            { name: 'Gravel', items: [] }
          ]
        },
        { name: 'Drains & Utilities', items: [] },
        { name: 'Alternate', items: [] }
      ]
    }
  ]
}

export default capstoneTemplate