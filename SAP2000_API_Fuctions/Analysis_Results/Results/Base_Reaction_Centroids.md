# Base Reaction Centroids

All base reaction data is considered in and reported in the global coordinate system. For clarity and to reduce wordiness, where the following description refers to a coordinate direction as X, Y or Z, it actually means global X, global Y or global Z.

**Important Note:** The base reaction centroids are calculated from translational reaction forces only. Reaction moments are not considered when calculating these centroids. Thus the reported base reaction centroids are not the same as the centroid of the applied loads.

Three centroid locations are reported: one for the X direction translational reactions, one for the Y direction translational reactions, and one for the Z direction translational reactions.

The centroid for the X direction translational reactions is (XCentroidForFx, YCentroidForFx, ZCentroidForFx).

The centroid for the Y direction translational reactions is (XCentroidForFy, YCentroidForFy, ZCentroidForFy).

The centroid for the Z direction translational reactions is (XCentroidForFz, YCentroidForFz, ZCentroidForFz).

Following are descriptions of each coordinate item.

XCentroidForFx

This is an array of the X coordinates of the centroid of all X-direction translational reaction forces for each result. [L]

This item is reported as zero if the sum of all translational reaction forces in the X- direction is zero.

YCentroidForFx

This is an array of the Y coordinates of the centroid of all X-direction translational reaction forces for each result. [L]

Said another way, the coordinate in this array is the Y coordinate about which the Mz moment caused by the X-direction translational reaction forces is zero.

This item is reported as zero if the sum of all translational reaction forces in the X-direction is zero.

ZCentroidForFx

This is an array of the Z coordinates of the centroid of all X-direction translational reaction forces for each result. [L]

Said another way, the coordinate in this array is the Z coordinate about which the My moment caused by the X-direction translational reaction forces is zero.

This item is reported as zero if the sum of all translational reaction forces in the X-direction is zero.

XCentroidForFy

This is an array of the X coordinates of the centroid of all Y-direction translational reaction forces for each result. [L]

Said another way, the coordinate in this array is the X coordinate about which the Mz moment caused by the Y-direction translational reaction forces is zero.

This item is reported as zero if the sum of all translational reaction forces in the Y-direction is zero.

YCentroidForFy

This is an array of the Y coordinates of the centroid of all Y-direction translational reaction forces for each result. [L]

This item is reported as zero if the sum of all translational reaction forces in the Y-direction is zero.

ZCentroidForFy

This is an array of the Z coordinates of the centroid of all Y-direction translational reaction forces for each result. [L]

Said another way, the coordinate in this array is the Z coordinate about which the Mx moment caused by the Y-direction translational reaction forces is zero.

This item is reported as zero if the sum of all translational reaction forces in the Y-direction is zero.

XCentroidForFz

This is an array of the X coordinates of the centroid of all Z-direction translational reaction forces for each result. [L]

Said another way, the coordinate in this array is the X coordinate about which the My moment caused by the Z-direction translational reaction forces is zero.

This item is reported as zero if the sum of all translational reaction forces in the Z-direction is zero.

YCentroidForFz

This is an array of the Y coordinates of the centroid of all Z-direction translational reaction forces for each result. [L]

Said another way, the coordinate in this array is the Y coordinate about which the Mx moment caused by the Z-direction translational reaction forces is zero.

This item is reported as zero if the sum of all translational reaction forces in the Z-direction is zero.

ZCentroidForFz

This is an array of the Z coordinates of the centroid of all Z-direction translational reaction forces for each result. [L]

This item is reported as zero if the sum of all translational reaction forces in the Z-direction is zero.
