use_container_width=True
)

    with right:
    
            st.subheader("☁️ 3. Informasi Publish")
            
            if publish:
                token = st.secrets["GITHUB_TOKEN"]
                repo = st.secrets["GITHUB_REPO"]
                branch = st.secrets["GITHUB_BRANCH"]
            
                with open(kmz_path, "rb") as file:
                    content = base64.b64encode(file.read()).decode()
            
                url = f"https://api.github.com/repos/{repo}/contents/ODP_Master.kmz"
            
                headers = {
                    "Authorization": f"Bearer {token}",
                    "Accept": "application/vnd.github+json"
                }
            
                get = requests.get(url, headers=headers)
            
                sha = None
            
                if get.status_code == 200:
                    sha = get.json()["sha"]
            
                payload = {
                    "message": f"Update KMZ {today}",
                    "content": content,
                    "branch": branch
                }
            
                if sha:
                    payload["sha"] = sha
            
                response = requests.put(
                    url,
                    headers=headers,
                    json=payload
                )
            
                if response.status_code in [200, 201]:
            
                    st.success("✔ Publish berhasil!")
            
                    st.markdown("""
            ### ☁️ 3. Informasi Publish
            """)
            
                    st.markdown(f"""
            | Informasi | Nilai |
            |-----------|-------|
            | 📅 Update | {today} |
            | 📍 Total ODP | {stats["total"]} |
            | ☁️ Status | GitHub berhasil diperbarui |
            | 🌿 Branch | {branch} |
            | 📂 Repository | {repo} |
            """)
            
                    st.markdown(f"""
            - 📅 **Update** : {today}
            - 📍 **Total ODP** : {stats["total"]}
            - ☁️ **Status** : GitHub berhasil diperbarui.
            """)
            
                else:
                    st.error("Publish gagal!")
            with right:

                    st.write(response.status_code)
                    st.write(response.json())
                    st.subheader("☁️ 3. Informasi Publish")
                    
                    if publish:
                        token = st.secrets["GITHUB_TOKEN"]
                        repo = st.secrets["GITHUB_REPO"]
                        branch = st.secrets["GITHUB_BRANCH"]
                    
                        with open(kmz_path, "rb") as file:
                            content = base64.b64encode(file.read()).decode()
                    
                        url = f"https://api.github.com/repos/{repo}/contents/ODP_Master.kmz"
                    
                        headers = {
                            "Authorization": f"Bearer {token}",
                            "Accept": "application/vnd.github+json"
                        }
                    
                        get = requests.get(url, headers=headers)
                    
                        sha = None
                    
                        if get.status_code == 200:
                            sha = get.json()["sha"]
                    
                        payload = {
                            "message": f"Update KMZ {today}",
                            "content": content,
                            "branch": branch
                        }
                    
                        if sha:
                            payload["sha"] = sha
                    
                        response = requests.put(
                            url,
                            headers=headers,
                            json=payload
                        )
                    
                        if response.status_code in [200, 201]:
                    
                            st.success("✔ Publish berhasil!")
                    
                            st.markdown("""
                    ### ☁️ 3. Informasi Publish
                    """)
                    
                            st.markdown(f"""
                    | Informasi | Nilai |
                    |-----------|-------|
                    | 📅 Update | {today} |
                    | 📍 Total ODP | {stats["total"]} |
                    | ☁️ Status | GitHub berhasil diperbarui |
                    | 🌿 Branch | {branch} |
                    | 📂 Repository | {repo} |
                    """)
                    
                            st.markdown(f"""
                    - 📅 **Update** : {today}
                    - 📍 **Total ODP** : {stats["total"]}
                    - ☁️ **Status** : GitHub berhasil diperbarui.
                    """)
                    
                        else:
                            st.error("Publish gagal!")
                    
                            st.write(response.status_code)
                            st.write(response.json())



