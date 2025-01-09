﻿// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { TypedUseSelectorHook, useDispatch, useSelector } from "react-redux";
import { configureStore } from "@reduxjs/toolkit";
import messagesReducer from "./messagesSlice";

export const store = configureStore({
    reducer: { messages: messagesReducer },
    devTools: process.env.NODE_ENV !== "production",
});

// Infer the `RootState` and `AppDispatch` types from the store itself
export type RootState = ReturnType<typeof store.getState>;

// Inferred type: {posts: PostsState, comments: CommentsState, users: UsersState}
export type AppDispatch = typeof store.dispatch;

export const useAppDispatch = () => useDispatch<AppDispatch>();
export const useAppSelector: TypedUseSelectorHook<RootState> = useSelector;
export enum TemplateSelection {
    Default = 'Default',
    infromational = 'Informational',
    infoVideo = 'Informational with Video',
    department = 'Department Message',
    departmentVideo = 'department message with poster',
    video = "Video",
    Default_ar = 'Default Arabic',
    infromational_ar = 'Informational Arabic',
    infoVideo_ar = 'Informational with Video Arabic',
    department_ar = 'Department Message Arabic',
    departmentVideo_ar = 'department message with poster Arabic',
    uae50 = "uae50 خمسون عام على الاتحاد",
}
