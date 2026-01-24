#pragma once
#ifdef __cplusplus
extern "C" {
#endif

typedef struct FFTContext FFTContext;

typedef struct FFTComplex {
    float re;
    float im;
} FFTComplex;

#ifdef __cplusplus
}
#endif
