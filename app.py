import streamlit as st
import logic
import os


# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="AI PPT Agent", page_icon="✨", layout="centered")


# --- SIDEBAR ---
with st.sidebar:
    st.header("Configuration")
    if logic.client:
        st.success("Groq Connection: Successful")
        st.info(f"Using Model: {logic.model_name}")
    else:
        st.error("Groq Connection: Failed")
        st.warning("Check your .env file for a valid GROQ_API_KEY.")
   
    st.divider()
    st.header("Presentation Settings")
    # Color Picker for Theme
    theme_color = st.color_picker("Pick a Theme Color", "#003366") # Default Navy Blue


# --- MAIN APP ---
st.title("✨ AI PowerPoint Creator Agent")
st.caption("Powered by Groq & Pexels")


# --- CHAT LOGIC ---
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content": "Hello! What topic would you like a presentation on?"}]
if "ppt_structure" not in st.session_state:
    st.session_state.ppt_structure = None


# Display chat history
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        if message.get("is_json"):
            st.json(message["content"])
        else:
            st.write(message["content"])


# Handle new user input
if prompt := st.chat_input("Type your message..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.write(prompt)


    with st.chat_message("assistant"):
        with st.spinner("Agent is thinking..."):
           
            # 1. Initial Generation
            if st.session_state.ppt_structure is None:
                response = logic.generate_slide_content(prompt)
               
                if response and "Error:" in response[:10]:
                    st.error(response)
                elif response:
                    st.session_state.ppt_structure = response
                    st.write("Here is the draft plan:")
                    st.json(response)
                    st.session_state.messages.append({"role": "assistant", "content": response, "is_json": True})
                   
                    follow_up_message = "Type 'Yes' to generate the file, or type feedback to change it."
                    st.write(follow_up_message)
                    st.session_state.messages.append({"role": "assistant", "content": follow_up_message})
                else:
                    st.error("Unknown error. Please check logs.")


            # 2. Feedback Loop
            else:
                if prompt.lower() in ["yes", "y", "looks good", "ok"]:
                    st.write(f"Generating images and building PowerPoint with theme {theme_color}... ⏳")
                   
                    # Pass the selected theme_color to the logic function
                    filename = logic.create_ppt_file(st.session_state.ppt_structure, include_images=True, theme_color=theme_color)
                   
                    if filename:
                        st.success("Presentation Ready!")
                        with open(filename, "rb") as file:
                            st.download_button(
                                label="📥 Download .pptx",
                                data=file,
                                file_name=filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        # Reset for next presentation
                        st.session_state.ppt_structure = None
                        st.session_state.messages.append({"role": "assistant", "content": "Great! What should we create next?"})
                    else:
                        st.error("Failed to create PPT file.")


                else: # Handle feedback
                    st.write("Refining content based on your feedback...")
                    response = logic.generate_slide_content(topic=prompt, feedback=prompt, current_content=st.session_state.ppt_structure)
                    if response and "Error:" not in response:
                        st.session_state.ppt_structure = response
                        st.write("Updated Draft:")
                        st.json(response)
                        st.session_state.messages.append({"role": "assistant", "content": response, "is_json": True})
                       
                        follow_up_message = "How is this? Type 'Yes' to generate."
                        st.write(follow_up_message)
                        st.session_state.messages.append({"role": "assistant", "content": follow_up_message})
                    else:
                        st.error(response)
